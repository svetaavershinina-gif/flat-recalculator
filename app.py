import streamlit as st
import pandas as pd
import io
import zipfile

st.set_page_config(page_title="Переоценка квартир", page_icon="🏠", layout="centered")
st.title("🏠 Переоценка квартир по проектам")
st.markdown("""
Выберите файл Excel с квартирами, статус, проекты и сумму прибавки.  
На выходе вы получите архив с переоценёнными файлами по каждому проекту.
""")

uploaded_file = st.file_uploader("📥 Загрузите Excel-файл", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Ошибка при чтении файла: {e}")
    else:
        required_cols = {"Статус", "Проект", "Стоимость"}
        if not required_cols.issubset(df.columns):
            st.error(f"Файл должен содержать колонки: {', '.join(required_cols)}")
        else:
            status_choice = st.radio("Что переоценить:", ["в строительстве", "сданные"], index=0)
            add_val = st.number_input("Сколько добавить к стоимости (₽):", step=10000, min_value=0)
            projects = df[df["Статус"] == status_choice]["Проект"].unique()
            chosen_projects = st.multiselect("Выберите проекты для переоценки", projects)

            if st.button("Выполнить пересчёт"):
                if not chosen_projects:
                    st.warning("Сначала выберите проекты!")
                else:
                    buffer = io.BytesIO()
                    with zipfile.ZipFile(buffer, "w") as zf:
                        for proj in chosen_projects:
                            df_proj = df[(df["Статус"] == status_choice) & (df["Проект"] == proj)].copy()
                            df_proj["Новая стоимость"] = df_proj["Стоимость"] + add_val
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