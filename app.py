import streamlit as st
import pandas as pd
import io
import zipfile

# Настройка страницы
st.set_page_config(
    page_title="Переоценка квартир",
    page_icon="my_icon.png",  # <- сюда можно вставить свой PNG или emoji "🏡"
    layout="wide"
)

st.title("🏠 Переоценка квартир по подразделениям")
st.markdown("""
Загрузите Excel с квартирами, выберите готовность объекта, подразделения и сумму прибавки.  
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
            # Выбор готовности объекта
            readiness_options = df["Готовность объекта"].unique()
            readiness_choice = st.selectbox("Выберите готовность объекта:", readiness_options)

            # Фильтруем по готовности
            df_filtered = df[df["Готовность объекта"] == readiness_choice]

            # Выбор подразделений
            department_options = df_filtered["Подразделение"].unique()
            chosen_departments = st.multiselect("Выберите подразделения для переоценки:", department_options)

            # Ввод суммы прибавки
            add_val = st.number_input("Сколько добавить к стоимости (₽):", step=10000, min_value=0)

            # Превью данных
            if chosen_departments:
                preview_df = df_filtered[df_filtered["Подразделение"].isin(chosen_departments)].copy()
                preview_df["Новая стоимость"] = preview_df["Стоимость"] + add_val
                preview_df = preview_df[["Номер квартиры", "Этаж", "Площадь общая",
                                         "Тип квартиры", "Стоимость", "Новая стоимость"]]
                st.subheader("Превью таблицы с новой стоимостью")
                st.dataframe(preview_df)

            # Генерация ZIP
            if st.button("Выполнить пересчёт"):
                if not chosen_departments:
                    st.warning("Сначала выберите хотя бы одно подразделение!")
                else:
                    buffer = io.BytesIO()
                    with zipfile.ZipFile(buffer, "w") as zf:
                        for dept in chosen_departments:
                            df_dept = df_filtered[df_filtered["Подразделение"] == dept].copy()
                            df_dept["Новая стоимость"] = df_dept["Стоимость"] + add_val
                            df_dept = df_dept[["Номер квартиры", "Этаж", "Площадь общая",
                                               "Тип квартиры", "Стоимость", "Новая стоимость"]]
                            out_file = io.BytesIO()
                            df_dept.to_excel(out_file, index=False)
                            zf.writestr(f"{dept}.xlsx", out_file.getvalue())
                    buffer.seek(0)
                    st.success("Файлы готовы! Скачайте архив ниже.")
                    st.download_button(
                        label="📥 Скачать архив",
                        data=buffer,
                        file_name="recalculated_departments.zip",
                        mime="application/zip"
                    )