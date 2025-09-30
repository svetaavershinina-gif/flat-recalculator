import streamlit as st
import pandas as pd
import io
import zipfile

st.set_page_config(page_title="Переоценка квартир", page_icon="🏠", layout="wide")

st.title("🏠 Переоценка квартир по проектам")
st.markdown("""
Загрузите Excel с квартирами, выберите готовность объекта, проекты и сумму прибавки.  
Приложение покажет превью данных и создаст архив с переоценёнными файлами.
""")
st.markdown("Версия 2.0 — 30.09.2025")

uploaded_file = st.file_uploader("📥 Загрузите Excel-файл", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Ошибка при чтении файла: {e}")
    else:
        required_cols = {"Готовность объекта", "Проект", "Номер квартиры", "Этаж",
                         "Площадь общая", "Тип квартиры", "Стоимость"}
        if not required_cols.issubset(df.columns):
            st.error(f"Файл должен содержать колонки: {', '.join(required_cols)}")
        else:
            # Выбор готовности объекта
            readiness_options = df["Готовность объекта"].unique()
            readiness_choice = st.selectbox("Выберите готовность объекта:", readiness_options)

            # Фильтруем по готовности
            df_filtered = df[df["Готовность объекта"] == readiness_choice]

            # Выбор проектов
            project_options = df_filtered["Проект"].unique()
            chosen_projects = st.multiselect("Выберите проекты для переоценки:", project_options)

            # Ввод суммы прибавки
            add_val = st.number_input("Сколько добавить к стоимости (₽):", step=10000, min_value=0)

            # Превью данных
            if chosen_projects:
                preview_df = df_filtered[df_filtered["Проект"].isin(chosen_projects)].copy()
                preview_df["Новая стоимость"] = preview_df["Стоимость"] + add_val
                preview_df = preview_df[["Номер квартиры", "Этаж", "Площадь общая",
                                         "Тип квартиры", "Стоимость", "Новая стоимость"]]
                st.subheader("Превью таблицы с новой стоимостью")
                st.dataframe(preview_df)

            # Генерация ZIP
            if st.button("Выполнить пересчёт"):
                if not chosen_projects:
                    st.warning("Сначала выберите хотя бы один проект!")
                else:
                    buffer = io.BytesIO()
                    with zipfile.ZipFile(buffer, "w") as zf:
                        for proj in chosen_projects:
                            df_proj = df_filtered[df_filtered["Проект"] == proj].copy()
                            df_proj["Новая стоимость"] = df_proj["Стоимость"] + add_val
                            df_proj = df_proj[["Номер квартиры", "Этаж", "Площадь общая",
                                               "Тип квартиры", "Стоимость", "Новая стоимость"]]
                            out_file = io.BytesIO()
                            df_proj.to_excel(out_file, index=False)
                            zf.writestr(f"{proj}.xlsx", out_file.getvalue())
                    buffer.seek(0)
                    st.success("Файлы готовы! Скачайте архив ниже.")
                    st.download_button(
                        label="📥 Скачать архив",
                        data=buffer,
                        file_name="recalculated_projects.zip",
                        mime="application/zip"
                    )
