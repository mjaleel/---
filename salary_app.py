import streamlit as st
import pandas as pd
from datetime import datetime
from num2words import num2words
import os
import csv
import re
import shutil
import math
import json
import io
import zipfile
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from itertools import groupby

# ====================================================================
# الثوابت والبيانات الثابتة
# ====================================================================

PAYER_NAME = "مديرية تربية البصرة"
PAYER_ACCOUNT = "IQ26RAFB002100366585001"
CURRENCY = "IQD"
DETAILS_OF_CHARGES = "SLEV"
REMITTANCE_INFO_TEMPLATE = "SALARY {} {}"
MAX_ROWS_PER_FILE = 4000
MAX_AMOUNT_PER_FILE = 4_000_000_000

BANK_BICS = {
    'RAFB': 'RAFBIQB1098', 'RDBA': 'RDBAIQB1046', 'AIBI': 'AIBIIQBA991',
    'IDBQ': 'IDBQIQBA004', 'AINI': 'AINIIQBA015', 'NBIQ': 'NBIQIQBA830'
}

ARABIC_BANK_NAME_MAP = {
    'RAFB': 'الرافدين', 'RDBA': 'الرشيد', 'AIBI': 'آشور',
    'IDBQ': 'التنمية', 'AINI': 'الطيف', 'NBIQ': 'الأهلي'
}

BANKS_WITH_DYNAMIC_BRANCHES = ['AINI', 'NBIQ']
BANK_KEYS_FOR_FILTERING = list(BANK_BICS.keys())

ARABIC_MONTHS = {
    1: "كانون الثاني", 2: "شباط", 3: "آذار", 4: "نيسان", 5: "أيار", 6: "حزيران",
    7: "تموز", 8: "آب", 9: "أيلول", 10: "تشرين الأول", 11: "تشرين الثاني", 12: "كانون الأول"
}

FINAL_EXCEL_COLS = [
    'Reference', 'Value Date', 'Payer Name', 'Payer Acount', 'Amount',
    'Currency', 'Receiver BIC', 'Beneficiary Name', 'Beneficiary Acount',
    'Remittance Information', 'Details of Charges'
]

BRANCHES_FILE = "custom_branches.json"
ARABIC_DIGITS = str.maketrans('0123456789', '٠١٢٣٤٥٦٧٨٩')

# ====================================================================
# إدارة الفروع
# ====================================================================

def load_custom_branches():
    if "custom_branches" not in st.session_state:
        if os.path.exists(BRANCHES_FILE):
            try:
                with open(BRANCHES_FILE, 'r', encoding='utf-8') as f:
                    st.session_state.custom_branches = json.load(f)
            except:
                st.session_state.custom_branches = {"NBIQ": ["NBIQIQBA830", "NBIQIQBA863"]}
        else:
            st.session_state.custom_branches = {"NBIQ": ["NBIQIQBA830", "NBIQIQBA863"]}
    return st.session_state.custom_branches


def save_custom_branches(branches_dict):
    st.session_state.custom_branches = branches_dict
    with open(BRANCHES_FILE, 'w', encoding='utf-8') as f:
        json.dump(branches_dict, f, ensure_ascii=False, indent=2)

# ====================================================================
# الدوال المساعدة
# ====================================================================

def convert_amount_to_arabic(amount):
    try:
        val = float(amount)
        text = num2words(int(val), lang='ar')
        text = text.replace("ألفاً", "ألفًا").replace("مليوناً", "مليونًا").replace("ملياراً", "مليارًا")
        text = text.replace(" و ", " و").strip()
        return text + " دينار لا غير"
    except:
        return "صفر دينار"


def format_number(value, use_arabic_digits=False):
    try:
        formatted = f"{int(value):,}"
        if use_arabic_digits:
            formatted = formatted.translate(ARABIC_DIGITS)
            formatted = formatted.replace(',', '،')
        return formatted
    except:
        return str(value)


def apply_financial_rounding(df, column_name):
    df = df.copy().reset_index(drop=False)
    original_total = round(df[column_name].sum(), 2)
    df['_floor_val'] = df[column_name].apply(math.floor)
    df['_decimal_part'] = df[column_name] - df['_floor_val']
    total_to_distribute = int(round(original_total - df['_floor_val'].sum()))
    df_sorted = df.sort_values(by='_decimal_part', ascending=False).copy()
    df_sorted['_final_amount'] = df_sorted['_floor_val']
    if total_to_distribute > 0:
        top_positions = df_sorted.index[:total_to_distribute]
        df_sorted.loc[top_positions, '_final_amount'] += 1
    df_sorted = df_sorted.sort_values(by='index').set_index('index')
    df_sorted.index.name = None
    df_sorted[column_name] = df_sorted['_final_amount'].astype(int)
    return df_sorted.drop(columns=['_floor_val', '_decimal_part', '_final_amount'])


def get_receiver_bic_dynamic(row, custom_branches):
    key = str(row['Bank Key']).strip().upper()
    iban = str(row['Iban']).strip().upper()
    if key in BANKS_WITH_DYNAMIC_BRANCHES:
        try:
            if len(iban) >= 11:
                branch_code = iban[8:11]
                computed_bic = f"{key}IQBA{branch_code}"
                return computed_bic
            return BANK_BICS.get(key, 'UnknownBIC')
        except:
            return BANK_BICS.get(key, 'UnknownBIC')
    return BANK_BICS.get(key, 'UnknownBIC')


def adjust_column_width(writer, sheet_name='Sheet1'):
    ws = writer.sheets[sheet_name]
    for i, col in enumerate(ws.columns):
        max_length = 0
        column = get_column_letter(i + 1)
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        ws.column_dimensions[column].width = max(max_length + 3, 15)


def guess_column(keywords, columns):
    for kw in keywords:
        for col in columns:
            if kw.lower() in str(col).lower():
                return col
    return columns[0] if columns else ""

# ====================================================================
# المعالجة الرئيسية — تُعيد dict من اسم_الملف: BytesIO
# ====================================================================

def process_excel_data(uploaded_file, col_name, col_iban, col_salary):
    custom_branches = load_custom_branches()
    logs = []

    df = pd.read_excel(uploaded_file)

    missing = [c for c in [col_name, col_iban, col_salary] if c not in df.columns]
    if missing:
        raise ValueError(f"الأعمدة التالية غير موجودة في الملف: {', '.join(missing)}")

    df = df.rename(columns={
        col_name:   'الاسم',
        col_iban:   'Iban',
        col_salary: 'الراتب الصافي'
    })

    df['الراتب الصافي'] = pd.to_numeric(df['الراتب الصافي'], errors='coerce').fillna(0)
    df = df[df['الراتب الصافي'] != 0].dropna(subset=['Iban', 'الاسم'])

    original_total = round(df['الراتب الصافي'].sum(), 2)
    decimal_count = int((df['الراتب الصافي'] % 1 != 0).sum())
    floor_total = df['الراتب الصافي'].apply(math.floor).sum()
    dinars_to_dist = int(round(original_total - floor_total))

    logs.append(f"الأعشار: {decimal_count} صف | المجموع: {original_total:,.2f} | دنانير للتوزيع: {dinars_to_dist}")

    df = apply_financial_rounding(df, 'الراتب الصافي')
    final_total = df['الراتب الصافي'].sum()
    remaining_dec = int((df['الراتب الصافي'] % 1 != 0).sum())

    logs.append(f"✔ الجبر المالي: {original_total:,.2f} → {final_total:,} | فرق: {final_total - original_total:+.2f} | أعشار متبقية: {remaining_dec}")

    df['الاسم'] = df['الاسم'].astype(str).str[:35]
    today = datetime.now()
    date_str = today.strftime('%Y%m%d')
    month_ar = ARABIC_MONTHS.get(today.month, "")

    df['Value Date'] = date_str
    df['Payer Name'] = PAYER_NAME
    df['Payer Acount'] = PAYER_ACCOUNT
    df['Amount'] = df['الراتب الصافي']
    df['Currency'] = CURRENCY
    df['Details of Charges'] = DETAILS_OF_CHARGES
    df['Beneficiary Name'] = df['الاسم']
    df['Beneficiary Acount'] = df['Iban']
    df['Remittance Information'] = REMITTANCE_INFO_TEMPLATE.format(today.year, month_ar)
    df['Bank Key'] = df['Iban'].str[4:8]

    df_filtered = df[df['Bank Key'].isin(BANK_KEYS_FOR_FILTERING)].copy()
    df_filtered['Receiver BIC'] = df_filtered.apply(
        lambda row: get_receiver_bic_dynamic(row, custom_branches), axis=1
    )
    df_filtered['Reference'] = date_str + ' ' + df_filtered['Iban'].astype(str)

    output_files = {}
    grouped = df_filtered.groupby('Bank Key')
    file_count = 0

    for bic, bank_df in grouped:
        bank_key = bic[:4]
        bank_ar = ARABIC_BANK_NAME_MAP.get(bank_key, 'مصرف')
        num_rows = len(bank_df)
        start_row, file_index = 0, 1

        while start_row < num_rows:
            end_row = min(start_row + MAX_ROWS_PER_FILE, num_rows)
            current_slice = bank_df.iloc[start_row:end_row]
            while current_slice['Amount'].sum() > MAX_AMOUNT_PER_FILE and len(current_slice) > 1:
                end_row -= 1
                current_slice = bank_df.iloc[start_row:end_row]

            filename = f"{bank_ar}_الملف_{file_index}_{bank_key}_{date_str}.xlsx"
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                current_slice[FINAL_EXCEL_COLS].to_excel(writer, index=False, sheet_name='Sheet1')
                adjust_column_width(writer)
            buf.seek(0)
            output_files[filename] = buf.read()
            file_count += 1
            start_row = end_row
            file_index += 1

    stats = {
        "decimal_count": decimal_count,
        "dinars_to_dist": dinars_to_dist,
        "original_total": original_total,
        "final_total": final_total,
        "remaining_dec": remaining_dec,
        "file_count": file_count,
        "col_name": col_name,
        "col_iban": col_iban,
        "col_salary": col_salary,
    }

    return output_files, stats, logs

# ====================================================================
# دالة الملخص — تُعيد BytesIO
# ====================================================================

def create_summary_file(excel_files_dict, use_arabic_digits=False):
    """
    excel_files_dict: dict of {filename: bytes}
    """
    rows_data = []
    for filename, content in excel_files_dict.items():
        df = pd.read_excel(io.BytesIO(content))
        parts = filename.split('_')
        bank_name = parts[0]
        branch_code = parts[3] if len(parts) > 3 else "---"
        total_amount = int(df['Amount'].sum()) if 'Amount' in df.columns else 0
        rows_data.append({
            'المصرف': bank_name,
            'اسم الملف': filename,
            'رمز الفرع': branch_code,
            'عدد الموظفين': len(df),
            'المبلغ الإجمالي': total_amount,
        })

    rows_data.sort(key=lambda x: x['المصرف'])

    COLOR_HEADER = "1F4E78"
    COLOR_BANK_SUBTOT = "2E75B6"
    COLOR_GRAND_TOTAL = "C00000"
    COLOR_ROW_ODD = "EBF3FB"
    COLOR_ROW_EVEN = "FFFFFF"
    COLOR_BANK_HEADER = "D6E4F0"

    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    medium_border = Border(
        left=Side(style='medium'), right=Side(style='medium'),
        top=Side(style='medium'), bottom=Side(style='medium')
    )
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    right_align = Alignment(horizontal='right', vertical='center', wrap_text=True)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        pd.DataFrame().to_excel(writer, sheet_name='الملخص الشامل', index=False)
        ws = writer.sheets['الملخص الشامل']

        col_widths = [18, 50, 12, 16, 22, 65]
        col_letters = ['A', 'B', 'C', 'D', 'E', 'F']
        for letter, width in zip(col_letters, col_widths):
            ws.column_dimensions[letter].width = width

        ws.row_dimensions[1].height = 30
        ws.row_dimensions[2].height = 22

        ws.merge_cells('A1:F1')
        title_cell = ws['A1']
        title_cell.value = f"ملخص الحسابات الشامل — {PAYER_NAME}"
        title_cell.font = Font(bold=True, size=14, color="FFFFFF", name="Arial")
        title_cell.fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type="solid")
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        title_cell.border = medium_border

        headers = ['المصرف', 'اسم الملف المرجعي', 'رمز الفرع', 'عدد الموظفين', 'المبلغ الإجمالي (د.ع)', 'المبلغ كتابةً']
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col_idx, value=header)
            cell.font = Font(bold=True, size=11, color="FFFFFF", name="Arial")
            cell.fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type="solid")
            cell.alignment = center_align
            cell.border = thin_border

        current_row = 3
        grand_total_employees = 0
        grand_total_amount = 0

        for bank_name, bank_rows in groupby(rows_data, key=lambda x: x['المصرف']):
            bank_rows_list = list(bank_rows)
            bank_employees = sum(r['عدد الموظفين'] for r in bank_rows_list)
            bank_amount = sum(r['المبلغ الإجمالي'] for r in bank_rows_list)

            ws.row_dimensions[current_row].height = 22
            ws.merge_cells(f'A{current_row}:F{current_row}')
            bank_title_cell = ws[f'A{current_row}']
            bank_title_cell.value = f"🏦  مصرف  {bank_name}"
            bank_title_cell.font = Font(bold=True, size=12, color="1F4E78", name="Arial")
            bank_title_cell.fill = PatternFill(start_color=COLOR_BANK_HEADER, end_color=COLOR_BANK_HEADER, fill_type="solid")
            bank_title_cell.alignment = Alignment(horizontal='right', vertical='center')
            bank_title_cell.border = medium_border
            current_row += 1

            for file_idx, row_data in enumerate(bank_rows_list):
                ws.row_dimensions[current_row].height = 18
                bg_color = COLOR_ROW_ODD if file_idx % 2 == 0 else COLOR_ROW_EVEN
                values = [
                    row_data['المصرف'],
                    row_data['اسم الملف'],
                    row_data['رمز الفرع'],
                    row_data['عدد الموظفين'],
                    row_data['المبلغ الإجمالي'],
                    convert_amount_to_arabic(row_data['المبلغ الإجمالي']),
                ]
                for col_idx, val in enumerate(values, 1):
                    cell = ws.cell(row=current_row, column=col_idx, value=val)
                    cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
                    cell.border = thin_border
                    cell.font = Font(size=10, name="Arial")
                    cell.alignment = center_align if col_idx != 6 else right_align
                    if col_idx == 5:
                        if use_arabic_digits:
                            cell.value = format_number(val, use_arabic_digits=True)
                            cell.alignment = center_align
                        else:
                            cell.number_format = '#,##0'
                current_row += 1

            ws.row_dimensions[current_row].height = 24
            subtot_labels = [
                f"إجمالي {bank_name}", "───", "───",
                bank_employees, bank_amount,
                convert_amount_to_arabic(bank_amount),
            ]
            for col_idx, val in enumerate(subtot_labels, 1):
                cell = ws.cell(row=current_row, column=col_idx, value=val)
                cell.font = Font(bold=True, size=11, color="FFFFFF", name="Arial")
                cell.fill = PatternFill(start_color=COLOR_BANK_SUBTOT, end_color=COLOR_BANK_SUBTOT, fill_type="solid")
                cell.border = medium_border
                cell.alignment = center_align if col_idx != 6 else right_align
                if col_idx == 5:
                    if use_arabic_digits:
                        cell.value = format_number(val, use_arabic_digits=True)
                    else:
                        cell.number_format = '#,##0'
            current_row += 1

            ws.row_dimensions[current_row].height = 8
            for col_idx in range(1, 7):
                ws.cell(row=current_row, column=col_idx).fill = PatternFill(
                    start_color="F0F0F0", end_color="F0F0F0", fill_type="solid"
                )
            current_row += 1

            grand_total_employees += bank_employees
            grand_total_amount += bank_amount

        ws.row_dimensions[current_row].height = 30
        grand_labels = [
            "✦ المجموع الكلي النهائي ✦", "───", "───",
            grand_total_employees, grand_total_amount,
            convert_amount_to_arabic(grand_total_amount),
        ]
        for col_idx, val in enumerate(grand_labels, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            cell.font = Font(bold=True, size=12, color="FFFFFF", name="Arial")
            cell.fill = PatternFill(start_color=COLOR_GRAND_TOTAL, end_color=COLOR_GRAND_TOTAL, fill_type="solid")
            cell.border = medium_border
            cell.alignment = center_align if col_idx != 6 else right_align
            if col_idx == 5:
                if use_arabic_digits:
                    cell.value = format_number(val, use_arabic_digits=True)
                else:
                    cell.number_format = '#,##0'
        current_row += 1

        ws.row_dimensions[current_row].height = 28
        ws.merge_cells(f'A{current_row}:F{current_row}')
        written_cell = ws[f'A{current_row}']
        written_cell.value = f"المبلغ الكلي كتابةً :  {convert_amount_to_arabic(grand_total_amount)}"
        written_cell.font = Font(bold=True, size=11, color="FFFFFF", name="Arial")
        written_cell.fill = PatternFill(start_color="7B2C2C", end_color="7B2C2C", fill_type="solid")
        written_cell.alignment = Alignment(horizontal='right', vertical='center')
        written_cell.border = medium_border

        ws.freeze_panes = 'A3'
        ws.sheet_view.rightToLeft = True

    buf.seek(0)
    summary_stats = {
        "file_count": len(excel_files_dict),
        "grand_total_employees": grand_total_employees,
        "grand_total_amount": grand_total_amount,
    }
    return buf.read(), summary_stats

# ====================================================================
# دالة التشفير — تُعيد dict من اسم: bytes
# ====================================================================

def batch_convert_excel_to_txt(excel_files_dict):
    """
    تحويل ملفات Excel إلى CSV بصيغة بنكية صحيحة.
    الصيغة الناتجة:
      - السطر الأول: أسماء الأعمدة مفصولة بـ |
      - باقي الأسطر: القيم مفصولة بـ |
      - Amount: رقم صحيح نظيف بدون فاصلة وبدون تنصيص
      - الترميز: UTF-8 بدون BOM
    """
    result = {}

    for filename, content in excel_files_dict.items():
        df = pd.read_excel(io.BytesIO(content), dtype=str)
        base = os.path.splitext(filename)[0]

        # التحقق من وجود الأعمدة المطلوبة
        missing_cols = [c for c in FINAL_EXCEL_COLS if c not in df.columns]
        if missing_cols:
            raise ValueError(
                f"الأعمدة التالية غير موجودة في {filename}: {', '.join(missing_cols)}"
            )

        # الكتابة سطراً سطراً في الذاكرة
        lines = []

        # السطر الأول: الـ Headers
        lines.append('|'.join(FINAL_EXCEL_COLS))

        # باقي الأسطر: البيانات
        for _, row in df[FINAL_EXCEL_COLS].iterrows():
            line_values = []
            for col in FINAL_EXCEL_COLS:
                raw_val = str(row[col]).strip()
                if col == 'Amount':
                    # إزالة أي فاصلة أو علامات تنصيص وتحويل لرقم صحيح نظيف
                    clean = raw_val.replace(',', '').replace('"', '')
                    try:
                        line_values.append(str(int(float(clean))))
                    except ValueError:
                        raise ValueError(
                            f"قيمة Amount غير صالحة [{raw_val}] في الملف: {filename}"
                        )
                else:
                    line_values.append(raw_val)
            lines.append('|'.join(line_values))

        # تحويل لـ bytes بترميز UTF-8 بدون BOM
        csv_content = '\n'.join(lines).encode('utf-8')

        result[base + ".csv"] = csv_content

    return result
# ====================================================================
# تهيئة session_state
# ====================================================================

def init_session():
    defaults = {
        "processed_files": {},
        "processing_done": False,
        "col_name": None,
        "col_iban": None,
        "col_salary": None,
        "process_stats": None,
        "uploaded_df_columns": [],
        "custom_branches": {},
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v
    load_custom_branches()

# ====================================================================
# الواجهة الرئيسية
# ====================================================================

def main():
    st.set_page_config(
        page_title="نظام رواتب تربية البصرة",
        page_icon="🏛",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # CSS مخصص RTL
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap');

    * { direction: rtl; font-family: 'Tajawal', sans-serif !important; }

    .stApp {
        background: linear-gradient(135deg, #0a0f1e 0%, #0d1b2a 50%, #0a1628 100%);
        min-height: 100vh;
    }

    /* هيدر رئيسي */
    .main-header {
        background: linear-gradient(135deg, #1a237e 0%, #0d47a1 60%, #1565c0 100%);
        border-radius: 16px;
        padding: 28px 36px;
        margin-bottom: 24px;
        border: 1px solid #1976d2;
        box-shadow: 0 8px 32px rgba(13,71,161,0.4), 0 0 60px rgba(25,118,210,0.15);
        text-align: center;
    }
    .main-header h1 {
        color: #FFD700 !important;
        font-size: 2rem !important;
        font-weight: 800 !important;
        margin: 0 0 8px 0 !important;
        text-shadow: 0 2px 8px rgba(0,0,0,0.5);
    }
    .main-header p {
        color: #90CAF9 !important;
        font-size: 1rem !important;
        margin: 0 !important;
    }

    /* بطاقات الأقسام */
    .section-card {
        background: rgba(255,255,255,0.04);
        border: 1px solid rgba(255,255,255,0.1);
        border-radius: 12px;
        padding: 20px 24px;
        margin-bottom: 20px;
        backdrop-filter: blur(10px);
    }
    .section-title {
        color: #90CAF9;
        font-size: 1.05rem;
        font-weight: 700;
        margin-bottom: 14px;
        padding-bottom: 8px;
        border-bottom: 2px solid #1976d2;
    }

    /* بطاقات الإحصاء */
    .stat-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 12px;
        margin-top: 12px;
    }
    .stat-card {
        background: rgba(25,118,210,0.15);
        border: 1px solid rgba(25,118,210,0.3);
        border-radius: 10px;
        padding: 14px;
        text-align: center;
    }
    .stat-value {
        color: #FFD700;
        font-size: 1.4rem;
        font-weight: 800;
    }
    .stat-label {
        color: #90CAF9;
        font-size: 0.8rem;
        margin-top: 4px;
    }

    /* شريط الشريط الجانبي */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0d1b2a 0%, #0a1628 100%) !important;
        border-left: 1px solid rgba(255,255,255,0.08) !important;
    }
    [data-testid="stSidebar"] * { color: #e0e0e0 !important; }

    /* تنسيق الأزرار */
    .stButton > button {
        border-radius: 10px !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
        transition: all 0.2s ease !important;
        border: none !important;
        padding: 12px 20px !important;
    }
    .stButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(0,0,0,0.3) !important;
    }

    /* تنسيق رسائل النجاح والخطأ */
    .stSuccess, .stError, .stWarning, .stInfo {
        border-radius: 10px !important;
    }

    /* تنسيق المدخلات */
    .stSelectbox label, .stFileUploader label, .stRadio label {
        color: #90CAF9 !important;
        font-weight: 600 !important;
    }

    /* عنوان الشريط الجانبي */
    .sidebar-section {
        background: rgba(25,118,210,0.1);
        border: 1px solid rgba(25,118,210,0.2);
        border-radius: 10px;
        padding: 14px;
        margin-bottom: 16px;
    }
    .sidebar-section-title {
        color: #FFD700;
        font-weight: 700;
        font-size: 0.95rem;
        margin-bottom: 10px;
    }

    /* بطاقة الملف */
    .file-card {
        background: rgba(46,125,50,0.1);
        border: 1px solid rgba(46,125,50,0.3);
        border-radius: 10px;
        padding: 12px 16px;
        margin: 6px 0;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }

    div[data-testid="stDownloadButton"] > button {
        background: linear-gradient(135deg, #1B5E20, #2E7D32) !important;
        color: white !important;
        width: 100% !important;
    }

    /* تمييز قسم التحذير */
    .warning-box {
        background: rgba(183,28,28,0.12);
        border: 1px solid rgba(183,28,28,0.35);
        border-radius: 10px;
        padding: 14px 18px;
        color: #ef9a9a;
        font-size: 0.9rem;
        line-height: 1.7;
    }

    /* لوج */
    .log-box {
        background: rgba(0,0,0,0.35);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 8px;
        padding: 12px 16px;
        font-size: 0.85rem;
        color: #80CBC4;
        font-family: monospace !important;
        line-height: 1.8;
        max-height: 160px;
        overflow-y: auto;
    }
    </style>
    """, unsafe_allow_html=True)

    init_session()

    # ── الهيدر ──────────────────────────────────────────────────────
    st.markdown("""
    <div class="main-header">
        <h1>🏛 نظام معالجة رواتب تربية البصرة</h1>
        <p>جبر مالي  •  تقسيم مصرفي  •  ملخص منسق  •  تشفير</p>
    </div>
    """, unsafe_allow_html=True)

    # ====================================================================
    # الشريط الجانبي — الإعدادات
    # ====================================================================
    with st.sidebar:
        st.markdown("### ⚙️ الإعدادات")

        # ── قسم إدارة الفروع ──
        st.markdown('<div class="sidebar-section">', unsafe_allow_html=True)
        st.markdown('<div class="sidebar-section-title">🏦 إدارة فروع المصارف الديناميكية</div>', unsafe_allow_html=True)

        branches = load_custom_branches()

        selected_bank_mgr = st.selectbox(
            "المصرف:", BANKS_WITH_DYNAMIC_BRANCHES, key="branch_bank_select"
        )
        current_branches = branches.get(selected_bank_mgr, [])

        if current_branches:
            st.markdown("**الفروع المحفوظة:**")
            for b in current_branches:
                st.code(b, language=None)
        else:
            st.info("لا توجد فروع مضافة")

        new_bic = st.text_input(
            "BIC الفرع الجديد:",
            placeholder="مثال: NBIQIQBA863",
            key="new_bic_input"
        )
        col_add, col_del = st.columns(2)
        with col_add:
            if st.button("➕ إضافة", use_container_width=True, key="btn_add_branch"):
                bic = new_bic.strip().upper()
                if not bic:
                    st.warning("أدخل BIC أولاً")
                elif len(bic) < 8:
                    st.warning("BIC يجب 8+ أحرف")
                else:
                    if selected_bank_mgr not in branches:
                        branches[selected_bank_mgr] = []
                    if bic in branches[selected_bank_mgr]:
                        st.info("موجود مسبقاً")
                    else:
                        branches[selected_bank_mgr].append(bic)
                        save_custom_branches(branches)
                        st.success(f"✔ تمت الإضافة: {bic}")
                        st.rerun()

        del_bic = st.text_input(
            "BIC للحذف:",
            placeholder="مثال: NBIQIQBA830",
            key="del_bic_input"
        )
        with col_del:
            if st.button("🗑 حذف", use_container_width=True, key="btn_del_branch"):
                bic = del_bic.strip().upper()
                if selected_bank_mgr in branches and bic in branches[selected_bank_mgr]:
                    branches[selected_bank_mgr].remove(bic)
                    save_custom_branches(branches)
                    st.success(f"✔ تم الحذف: {bic}")
                    st.rerun()
                else:
                    st.warning("الفرع غير موجود")

        st.caption("⚠ BIC يُحسب تلقائياً من IBAN. الفروع للتوثيق فقط.")
        st.markdown('</div>', unsafe_allow_html=True)

        st.divider()

        # ── خيار الأرقام ──
        st.markdown('<div class="sidebar-section">', unsafe_allow_html=True)
        st.markdown('<div class="sidebar-section-title">🔢 نوع الأرقام في الملخص</div>', unsafe_allow_html=True)
        digit_mode = st.radio(
            "اختر:",
            ["أرقام إنجليزية  1,000", "أرقام عربية  ١،٠٠٠"],
            key="digit_mode_radio"
        )
        use_arabic_digits = digit_mode.startswith("أرقام عربية")
        st.markdown('</div>', unsafe_allow_html=True)

    # ====================================================================
    # المحتوى الرئيسي
    # ====================================================================
    tab1, tab2, tab3 = st.tabs(["① رفع الملف والمعالجة", "② الملخص والتشفير", "③ التنزيل"])

    # ──────────────────────────────────────────────────────────────────
    # تاب ١ — رفع الملف
    # ──────────────────────────────────────────────────────────────────
    with tab1:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">📂 رفع ملف Excel</div>', unsafe_allow_html=True)

        uploaded = st.file_uploader(
            "اختر ملف Excel الأصلي:",
            type=["xlsx"],
            key="main_uploader",
            help="ملف يحتوي على الاسم والـ IBAN والراتب الصافي"
        )
        st.markdown('</div>', unsafe_allow_html=True)

        if uploaded:
            try:
                df_preview = pd.read_excel(uploaded, nrows=5)
                columns = list(df_preview.columns)
                st.session_state.uploaded_df_columns = columns

                st.markdown('<div class="section-card">', unsafe_allow_html=True)
                st.markdown('<div class="section-title">🗂 تعيين أعمدة الملف</div>', unsafe_allow_html=True)
                st.caption("النظام يقترح الأنسب — راجع وعدّل عند الحاجة")

                default_name = guess_column(['اسم', 'name', 'الاسم', 'موظف'], columns)
                default_iban = guess_column(['iban', 'حساب', 'account', 'رقم'], columns)
                default_sal  = guess_column(['راتب', 'salary', 'صافي', 'net', 'مبلغ', 'amount'], columns)

                c1, c2, c3 = st.columns(3)
                with c1:
                    col_name = st.selectbox("👤 عمود الاسم:", columns,
                        index=columns.index(default_name) if default_name in columns else 0,
                        key="col_name_sel")
                with c2:
                    col_iban = st.selectbox("🏦 عمود IBAN:", columns,
                        index=columns.index(default_iban) if default_iban in columns else 0,
                        key="col_iban_sel")
                with c3:
                    col_salary = st.selectbox("💰 عمود الراتب الصافي:", columns,
                        index=columns.index(default_sal) if default_sal in columns else 0,
                        key="col_salary_sel")

                # تحقق من عدم تكرار الأعمدة
                selected_cols = [col_name, col_iban, col_salary]
                if len(set(selected_cols)) < 3:
                    st.error("⚠ اخترت نفس العمود لأكثر من حقل! تأكد من اختيار أعمدة مختلفة.")
                else:
                    st.session_state.col_name   = col_name
                    st.session_state.col_iban   = col_iban
                    st.session_state.col_salary = col_salary

                st.markdown('</div>', unsafe_allow_html=True)

                # معاينة
                st.markdown('<div class="section-card">', unsafe_allow_html=True)
                st.markdown('<div class="section-title">👁 معاينة أول 5 صفوف</div>', unsafe_allow_html=True)
                st.dataframe(df_preview, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

                # زر المعالجة
                if len(set(selected_cols)) == 3:
                    st.markdown('<div class="section-card">', unsafe_allow_html=True)
                    st.markdown('<div class="section-title">🚀 تنفيذ التقسيم والجبر المالي</div>', unsafe_allow_html=True)

                    if st.button("⑥ بدء التقسيم والمعالجة (جبر مالي) 🚀",
                                 use_container_width=True, type="primary", key="btn_process"):
                        with st.spinner("جاري المعالجة والجبر المالي..."):
                            try:
                                uploaded.seek(0)
                                output_files, stats, logs = process_excel_data(
                                    uploaded,
                                    st.session_state.col_name,
                                    st.session_state.col_iban,
                                    st.session_state.col_salary
                                )
                                st.session_state.processed_files = output_files
                                st.session_state.process_stats   = stats
                                st.session_state.processing_done = True

                                # عرض اللوج
                                log_html = "<br>".join(logs)
                                st.markdown(f'<div class="log-box">{log_html}</div>', unsafe_allow_html=True)

                                # إحصائيات
                                s = stats
                                st.markdown(f"""
                                <div class="stat-grid">
                                    <div class="stat-card">
                                        <div class="stat-value">{s['file_count']}</div>
                                        <div class="stat-label">عدد الملفات المنشأة</div>
                                    </div>
                                    <div class="stat-card">
                                        <div class="stat-value">{s['decimal_count']}</div>
                                        <div class="stat-label">صفوف ذات أعشار</div>
                                    </div>
                                    <div class="stat-card">
                                        <div class="stat-value">{s['dinars_to_dist']}</div>
                                        <div class="stat-label">دنانير موزعة</div>
                                    </div>
                                    <div class="stat-card">
                                        <div class="stat-value">{s['original_total']:,.2f}</div>
                                        <div class="stat-label">المجموع الأصلي</div>
                                    </div>
                                    <div class="stat-card">
                                        <div class="stat-value">{s['final_total']:,}</div>
                                        <div class="stat-label">المجموع بعد الجبر</div>
                                    </div>
                                    <div class="stat-card">
                                        <div class="stat-value" style="color:#4CAF50">{s['remaining_dec']}</div>
                                        <div class="stat-label">أعشار متبقية ✔</div>
                                    </div>
                                </div>
                                """, unsafe_allow_html=True)

                                st.success(f"✔ نجاح! تم إنشاء {s['file_count']} ملف — انتقل لتاب التنزيل")

                            except Exception as e:
                                st.error(f"خطأ: {e}")

                    st.markdown('</div>', unsafe_allow_html=True)

            except Exception as e:
                st.error(f"تعذّر قراءة الملف: {e}")
        else:
            st.info("⬆ ارفع ملف Excel لبدء العمل")

    # ──────────────────────────────────────────────────────────────────
    # تاب ٢ — الملخص والتشفير
    # ──────────────────────────────────────────────────────────────────
    with tab2:
        if not st.session_state.processing_done or not st.session_state.processed_files:
            st.warning("⚠ يجب إتمام المعالجة في التاب الأول أولاً")
        else:
            files_dict = st.session_state.processed_files

            # ── ملخص الحسابات ──
            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">📊 إنشاء ملخص الحسابات المفصّل</div>', unsafe_allow_html=True)

            if st.button("⑦ إنشاء ملخص الحسابات المفصّل 📊",
                         use_container_width=True, key="btn_summary"):
                with st.spinner("جاري إنشاء الملخص..."):
                    try:
                        summary_bytes, sum_stats = create_summary_file(files_dict, use_arabic_digits)
                        st.session_state["summary_bytes"] = summary_bytes
                        st.session_state["sum_stats"]     = sum_stats

                        ss = sum_stats
                        st.markdown(f"""
                        <div class="stat-grid">
                            <div class="stat-card">
                                <div class="stat-value">{ss['file_count']}</div>
                                <div class="stat-label">عدد الملفات</div>
                            </div>
                            <div class="stat-card">
                                <div class="stat-value">{ss['grand_total_employees']:,}</div>
                                <div class="stat-label">إجمالي الموظفين</div>
                            </div>
                            <div class="stat-card">
                                <div class="stat-value">{ss['grand_total_amount']:,}</div>
                                <div class="stat-label">المجموع الكلي (د.ع)</div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                        st.success("✔ تم إنشاء الملخص — يمكن تنزيله من تاب التنزيل")
                    except Exception as e:
                        st.error(f"خطأ بالملخص: {e}")
            st.markdown('</div>', unsafe_allow_html=True)

            # ── التشفير ──
            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">🔑 تشفير وتحويل الملفات (TXT/CSV)</div>', unsafe_allow_html=True)

            if st.button("⑧ تشفير وتحويل الملفات 🔑",
                         use_container_width=True, key="btn_encrypt"):
                with st.spinner("جاري التشفير والتحويل..."):
                    try:
                        enc_files = batch_convert_excel_to_txt(files_dict)
                        st.session_state["encrypted_files"] = enc_files
                        st.success(f"✔ تم تشفير {len(files_dict)} ملف — يمكن تنزيلها من تاب التنزيل")
                    except Exception as e:
                        st.error(f"خطأ بالتشفير: {e}")
            st.markdown('</div>', unsafe_allow_html=True)

    # ──────────────────────────────────────────────────────────────────
    # تاب ٣ — التنزيل
    # ──────────────────────────────────────────────────────────────────
    with tab3:
        if not st.session_state.processing_done:
            st.warning("⚠ يجب إتمام المعالجة أولاً")
        else:
            files_dict = st.session_state.processed_files

            # ── تنزيل الملفات المقسمة ──
            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">📥 تنزيل الملفات المقسمة (Excel)</div>', unsafe_allow_html=True)

            # زر تنزيل ZIP الكلي
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                for fname, fbytes in files_dict.items():
                    zf.writestr(fname, fbytes)
            zip_buf.seek(0)
            st.download_button(
                "📦 تنزيل جميع الملفات المقسمة (ZIP)",
                data=zip_buf.read(),
                file_name=f"salary_files_{datetime.now().strftime('%Y%m%d')}.zip",
                mime="application/zip",
                use_container_width=True,
                key="dl_all_zip"
            )

            st.markdown("---")
            st.markdown("**تنزيل كل ملف على حدة:**")
            for fname, fbytes in files_dict.items():
                c1, c2 = st.columns([3, 1])
                with c1:
                    st.markdown(f"📄 `{fname}`")
                with c2:
                    st.download_button(
                        "⬇ تنزيل",
                        data=fbytes,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{fname}"
                    )
            st.markdown('</div>', unsafe_allow_html=True)

            # ── تنزيل الملخص ──
            if "summary_bytes" in st.session_state:
                st.markdown('<div class="section-card">', unsafe_allow_html=True)
                st.markdown('<div class="section-title">📊 تنزيل ملخص الحسابات</div>', unsafe_allow_html=True)
                st.download_button(
                    "⬇ تنزيل ملف الملخص الشامل",
                    data=st.session_state["summary_bytes"],
                    file_name=f"Summary_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="dl_summary"
                )
                st.markdown('</div>', unsafe_allow_html=True)

            # ── تنزيل ملفات التشفير ──
            if "encrypted_files" in st.session_state:
                st.markdown('<div class="section-card">', unsafe_allow_html=True)
                st.markdown('<div class="section-title">🔑 تنزيل ملفات التشفير (TXT/CSV)</div>', unsafe_allow_html=True)

                enc_zip_buf = io.BytesIO()
                with zipfile.ZipFile(enc_zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for fname, fbytes in st.session_state["encrypted_files"].items():
                        zf.writestr(fname, fbytes)
                enc_zip_buf.seek(0)
                st.download_button(
                    "📦 تنزيل جميع ملفات TXT/CSV (ZIP)",
                    data=enc_zip_buf.read(),
                    file_name=f"encrypted_files_{datetime.now().strftime('%Y%m%d')}.zip",
                    mime="application/zip",
                    use_container_width=True,
                    key="dl_enc_zip"
                )

                st.markdown("---")
                for fname, fbytes in st.session_state["encrypted_files"].items():
                    c1, c2 = st.columns([3, 1])
                    with c1:
                        st.markdown(f"🔑 `{fname}`")
                    with c2:
                        mime = "text/csv"
                        st.download_button(
                            "⬇",
                            data=fbytes,
                            file_name=fname,
                            mime=mime,
                            key=f"dl_enc_{fname}"
                        )
                st.markdown('</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
