import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

# Настройка страницы
st.set_page_config(
    page_title="Переоценка квартир",
    page_icon="my_icon.png",  # <- сюда можно вставить свой PNG или emoji "🏡"
    layout="wide"
)

# Заголовок с картинкой
col1, col2 = st.columns([1, 10])
with col1:
    st.image("my_icon.png", width=50)
with col2:
    st.markdown("## Переоценка квартир по подразделениям")

st.markdown("""
Загрузите Excel с квартирами, выберите готовность объекта, подразделения, статус и сумму прибавки.  
Приложение покажет превью данных и создаст архив с переоценёнными файлами.
""")

# Загрузка файла
uploaded_file = st.file_uploader("📥 Загрузите Excel-файл", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Ошибка при чтении файла: {e}")
    else:
        required_cols = {"Готовность объекта", "Подразделение", "Номер квартиры", "Этаж",
                         "Площадь общая", "Тип квартиры", "Стоимость"}
        if not required_cols.issubset(df.columns):
            st.error(f"Файл должен содержать колонки: {', '.join(required_cols)}")
        else:
            # Фильтр по готовности
            readiness_options = df["Готовность объекта"].unique()
            readiness_choices = st.multiselect("Выберите готовность объекта:", readiness_options)
            df_filtered = df[df["Готовность объекта"].isin(readiness_choices)]

            # Фильтр по подразделению
            department_options = df_filtered["Подразделение"].unique()
            chosen_departments = st.multiselect("Выберите подразделения для переоценки:", department_options)

            # Фильтр по статусу
            if "Статус" in df_filtered.columns:
                status_options = df_filtered["Статус"].unique()
                chosen_status = st.multiselect("Выберите статус:", status_options)
                if chosen_status:
                    df_filtered = df_filtered[df_filtered["Статус"].isin(chosen_status)]
            else:
                st.warning("Колонка 'Статус' не найдена в файле")

            # Ввод суммы прибавки
            add_val = st.number_input("Сколько добавить к стоимости (₽):", step=10000, min_value=0)

            # Превью данных
            if readiness_choices and chosen_departments:
                preview_df = df_filtered[df_filtered["Подразделение"].isin(chosen_departments)].copy()
                preview_df["Новая стоимость"] = preview_df["Стоимость"] + add_val
                preview_df["Новая цена кв.м"] = preview_df["Новая стоимость"] / preview_df["Площадь общая"]
                preview_df["Изменение"] = preview_df["Новая стоимость"] - preview_df["Стоимость"]

                # Итоговая строка
                total_row = {
                    "Номер квартиры": "Итого",
                    "Этаж": "",
                    "Площадь общая": "",
                    "Тип квартиры": "",
                    "Стоимость": preview_df["Стоимость"].sum(),
                    "Новая стоимость": preview_df["Новая стоимость"].sum(),
                    "Новая цена кв.м": "",
                    "Изменение": preview_df["Изменение"].sum()
                }
                preview_df = pd.concat([preview_df, pd.DataFrame([total_row])], ignore_index=True)

                preview_df = preview_df[["Номер квартиры", "Этаж", "Площадь общая",
                                         "Тип квартиры", "Стоимость", "Новая стоимость",
                                         "Новая цена кв.м", "Изменение"]]
                st.subheader("Превью таблицы с новой стоимостью")
                st.dataframe(preview_df)

            # Генерация ZIP
            if st.button("Выполнить пересчёт"):
                if not readiness_choices:
                    st.warning("Сначала выберите хотя бы одну готовность объекта!")
                elif not chosen_departments:
                    st.warning("Сначала выберите хотя бы одно подразделение!")
                else:
                    buffer = io.BytesIO()
                    with zipfile.ZipFile(buffer, "w") as zf:
                        for dept in chosen_departments:
                            df_dept = df_filtered[df_filtered["Подразделение"] == dept].copy()
                            df_dept["Новая стоимость"] = df_dept["Стоимость"] + add_val
                            df_dept["Новая цена кв.м"] = df_dept["Новая стоимость"] / df_dept["Площадь общая"]
                            df_dept["Изменение"] = df_dept["Новая стоимость"] - df_dept["Стоимость"]

                            # Итоговая строка
                            total_row_dept = {
                                "Номер квартиры": "Итого",
                                "Этаж": "",
                                "Площадь общая": "",
                                "Тип квартиры": "",
                                "Стоимость": df_dept["Стоимость"].sum(),
                                "Новая стоимость": df_dept["Новая стоимость"].sum(),
                                "Новая цена кв.м": "",
                                "Изменение": df_dept["Изменение"].sum()
                            }
                            df_dept = pd.concat([df_dept, pd.DataFrame([total_row_dept])], ignore_index=True)
                            df_dept = df_dept[["Номер квартиры", "Этаж", "Площадь общая",
                                               "Тип квартиры", "Стоимость", "Новая стоимость",
                                               "Новая цена кв.м", "Изменение"]]

                            # Сохраняем в Excel и форматируем границы только вокруг таблицы
                            out_file = io.BytesIO()
                            df_dept.to_excel(out_file, index=False, engine='openpyxl')
                            out_file.seek(0)
                            wb = load_workbook(out_file)
                            ws = wb.active

                            # Применяем границы только к диапазону с данными
                            max_row = ws.max_row
                            max_col = ws.max_column
                            thin = Side(border_style="thin", color="000000")
                            border = Border(left=thin, right=thin, top=thin, bottom=thin)
                            for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
                                for cell in row:
                                    cell.border = border

                            # Сохраняем обратно в BytesIO и пишем в ZIP
                            final_buffer = io.BytesIO()
                            wb.save(final_buffer)
                            final_buffer.seek(0)
                            zf.writestr(f"{dept}.xlsx", final_buffer.getvalue())

                    buffer.seek(0)
                    st.success("Файлы готовы! Скачайте архив ниже.")
                    st.download_button(
                        label="📥 Скачать архив",
                        data=buffer,
                        file_name="recalculated_departments.zip",
                        mime="application/zip"
                    )