import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm

# --- Функция для умного PDF ---
def excel_to_pdf_smart(df, title):
    buffer = io.BytesIO()
    pdf = SimpleDocTemplate(buffer, pagesize=A4,
                            topMargin=2*cm, bottomMargin=2*cm,
                            leftMargin=2*cm, rightMargin=2*cm)
    elements = []

    data = [df.columns.tolist()] + df.values.tolist()

    # Настройки шрифта и высоты строки
    font_size = 8
    row_height = font_size * 1.8  # высота строки с padding

    # Доступная высота для таблицы
    available_height = A4[1] - 4*cm  # верхние и нижние поля

    max_rows_per_page = int(available_height // row_height)
    if max_rows_per_page < 1:
        max_rows_per_page = 20

    # Масштабируем ширину колонок по странице
    n_cols = len(df.columns)
    page_width = A4[0] - 4*cm
    col_width = page_width / n_cols
    col_widths = [col_width] * n_cols

    # Разбиваем таблицу на страницы
    for i in range(0, len(data)-1, max_rows_per_page):
        chunk = [data[0]] + data[i+1:i+1+max_rows_per_page]
        table = Table(chunk, repeatRows=1, colWidths=col_widths)
        style = TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F79646')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), font_size),
            ('BOTTOMPADDING', (0,0), (-1,0), 6),
            ('GRID', (0,0), (-1,-1), 0.25, colors.black),
        ])
        table.setStyle(style)
        elements.append(table)
        elements.append(PageBreak())

    pdf.build(elements)
    buffer.seek(0)
    return buffer

# --- Настройка страницы ---
st.set_page_config(
    page_title="Переоценка квартир",
    page_icon="🏡",
    layout="wide"
)

col1, col2 = st.columns([1, 10])
with col1:
    st.image("my_icon.png", width=50)
with col2:
    st.markdown("## Переоценка квартир по подразделениям")

st.markdown("""
Загрузите Excel с квартирами, выберите фильтры и сумму прибавки.  
Приложение покажет превью данных и создаст архив с переоценёнными файлами.
""")

# --- Загрузка файла ---
uploaded_file = st.file_uploader("📥 Загрузите Excel-файл", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Ошибка при чтении файла: {e}")
    else:
        required_cols = {"Готовность объекта", "Подразделение", "Номер квартиры", "Этаж",
                         "Площадь общая", "Тип квартиры", "Статус", "Вид помещения", "Стоимость"}
        if not required_cols.issubset(df.columns):
            st.error(f"Файл должен содержать колонки: {', '.join(required_cols)}")
        else:
            # --- Фильтры ---
            readiness_options = df["Готовность объекта"].unique()
            readiness_choices = st.multiselect("Выберите готовность объекта:", readiness_options)
            df_filtered = df[df["Готовность объекта"].isin(readiness_choices)] if readiness_choices else df.copy()

            department_options = df_filtered["Подразделение"].unique()
            chosen_departments = st.multiselect("Выберите подразделения для переоценки:", department_options)
            if chosen_departments:
                df_filtered = df_filtered[df_filtered["Подразделение"].isin(chosen_departments)]

            status_options = df_filtered["Статус"].unique()
            chosen_status = st.multiselect("Выберите статус:", status_options)
            if chosen_status:
                df_filtered = df_filtered[df_filtered["Статус"].isin(chosen_status)]

            vid_options = df_filtered["Вид помещения"].unique()
            chosen_vid = st.multiselect("Выберите вид помещения:", vid_options)
            if chosen_vid:
                df_filtered = df_filtered[df_filtered["Вид помещения"].isin(chosen_vid)]

            flat_type_options = df_filtered["Тип квартиры"].unique()
            chosen_flat_types = st.multiselect("Выберите тип квартиры:", flat_type_options)
            if chosen_flat_types:
                df_filtered = df_filtered[df_filtered["Тип квартиры"].isin(chosen_flat_types)]

            add_val = st.number_input("Сколько добавить к стоимости (₽):", step=10000, min_value=0)
            report_date = st.date_input("📅 Выберите дату для имени файла")

            # --- Превью данных ---
            if not df_filtered.empty and chosen_departments:
                preview_df = df_filtered.copy()
                preview_df["Новая стоимость"] = preview_df["Стоимость"] + add_val
                preview_df["Новая цена кв.м"] = preview_df["Новая стоимость"] / preview_df["Площадь общая"]
                preview_df["Изменение"] = preview_df["Новая стоимость"] - preview_df["Стоимость"]

                preview_df = preview_df[[
                    "Номер квартиры", "Этаж", "Площадь общая", "Тип квартиры",
                    "Статус", "Вид помещения", "Стоимость", "Новая стоимость",
                    "Новая цена кв.м", "Изменение"
                ]]

                totals = {
                    "Номер квартиры": "ИТОГО:",
                    "Стоимость": preview_df["Стоимость"].sum(),
                    "Новая стоимость": preview_df["Новая стоимость"].sum(),
                    "Изменение": preview_df["Изменение"].sum()
                }
                totals_df = pd.DataFrame([totals])
                preview_with_total = pd.concat([preview_df, totals_df], ignore_index=True)
                st.subheader("Превью таблицы с новой стоимостью")
                st.dataframe(preview_with_total)

            # --- Генерация ZIP ---
            if st.button("Выполнить пересчёт"):
                if df_filtered.empty:
                    st.warning("Нет данных для пересчёта по выбранным фильтрам!")
                else:
                    date_str = report_date.strftime("%d.%m.%Y")
                    excel_buffer = io.BytesIO()
                    pdf_buffer = io.BytesIO()

                    with zipfile.ZipFile(excel_buffer, "w") as zf_excel, zipfile.ZipFile(pdf_buffer, "w") as zf_pdf:
                        for dept in chosen_departments:
                            df_dept = df_filtered[df_filtered["Подразделение"] == dept].copy()
                            df_dept["Новая стоимость"] = df_dept["Стоимость"] + add_val
                            df_dept["Новая цена кв.м"] = df_dept["Новая стоимость"] / df_dept["Площадь общая"]
                            df_dept["Изменение"] = df_dept["Новая стоимость"] - df_dept["Стоимость"]

                            df_dept = df_dept[[
                                "Номер квартиры", "Этаж", "Площадь общая", "Тип квартиры",
                                "Статус", "Вид помещения", "Стоимость", "Новая стоимость",
                                "Новая цена кв.м", "Изменение"
                            ]]

                            totals = {
                                "Номер квартиры": "ИТОГО:",
                                "Стоимость": df_dept["Стоимость"].sum(),
                                "Новая стоимость": df_dept["Новая стоимость"].sum(),
                                "Изменение": df_dept["Изменение"].sum()
                            }
                            totals_df = pd.DataFrame([totals])
                            df_dept = pd.concat([df_dept, totals_df], ignore_index=True)

                            # --- Excel ---
                            wb = Workbook()
                            ws = wb.active
                            for r in dataframe_to_rows(df_dept, index=False, header=True):
                                ws.append(r)

                            thin = Side(border_style="thin", color="000000")
                            border = Border(left=thin, right=thin, top=thin, bottom=thin)
                            header_fill = PatternFill("solid", fgColor="F79646")
                            zebra_fill = PatternFill("solid", fgColor="D9D9D9")
                            header_font = Font(name="Calibri Light", size=9, bold=True)
                            cell_font = Font(name="Calibri Light", size=9)
                            align_center = Alignment(horizontal="center", vertical="center")

                            max_row = ws.max_row
                            max_col = ws.max_column

                            for cell in ws[1]:
                                cell.fill = header_fill
                                cell.font = header_font
                                cell.alignment = align_center
                                cell.border = border

                            for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=max_row-1, min_col=1, max_col=max_col), start=1):
                                for cell in row:
                                    cell.font = cell_font
                                    cell.border = border
                                    cell.alignment = align_center
                                    if row_idx % 2 == 0:
                                        cell.fill = zebra_fill
                                    if cell.column == 3:
                                        cell.number_format = '#,##0.00'
                                    elif isinstance(cell.value, (int, float)):
                                        cell.number_format = '#,##0'

                            for cell in ws[max_row]:
                                cell.fill = header_fill
                                cell.font = header_font
                                cell.alignment = align_center
                                cell.border = border
                                if cell.column == 3:
                                    cell.number_format = '#,##0.00'
                                elif isinstance(cell.value, (int, float)):
                                    cell.number_format = '#,##0'

                            for col in ws.columns:
                                max_length = 0
                                col_letter = col[0].column_letter
                                for cell in col:
                                    if cell.value:
                                        max_length = max(max_length, len(str(cell.value)))
                                ws.column_dimensions[col_letter].width = max_length + 2

                            final_buffer = io.BytesIO()
                            wb.save(final_buffer)
                            final_buffer.seek(0)
                            zf_excel.writestr(f"{dept}_{date_str}.xlsx", final_buffer.getvalue())

                            # --- PDF ---
                            pdf_file = excel_to_pdf_smart(df_dept, dept)
                            zf_pdf.writestr(f"{dept}_{date_str}.pdf", pdf_file.getvalue())

                    excel_buffer.seek(0)
                    pdf_buffer.seek(0)

                    st.success("Файлы готовы! Скачайте архивы ниже.")
                    st.download_button(
                        label="📥 Скачать архив Excel",
                        data=excel_buffer,
                        file_name="recalculated_departments.zip",
                        mime="application/zip"
                    )
                    st.download_button(
                        label="📥 Скачать архив PDF",
                        data=pdf_buffer,
                        file_name="recalculated_departments_pdf.zip",
                        mime="application/zip"
                    )