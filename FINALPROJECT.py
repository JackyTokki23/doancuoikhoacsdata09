import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os, shutil

st.set_page_config(page_title="CSV Report Generator", layout="wide")
st.title("CSV Report Generator")

f = st.file_uploader("Upload CSV", type=["csv"])

def clean_df(df):
    for c in df.columns:
        df[c] = df[c].replace(["?", "NA", "N/A", "nan", "null", "--", ""], pd.NA)
        if df[c].dtype == object:
            try:
                df[c] = pd.to_datetime(df[c], errors="ignore")
            except:
                pass
            try:
                df[c] = pd.to_numeric(df[c].astype(str).str.replace(",", ""), errors="ignore")
            except:
                pass
    return df

def make_report(df, name="report.xlsx"):
    charts = "temp_charts"
    os.makedirs(charts, exist_ok=True)

    with pd.ExcelWriter(name, engine="xlsxwriter") as w:
        info = {"Rows":[len(df)],"Cols":[len(df.columns)],"Names":[", ".join(df.columns)]}
        pd.DataFrame(info).to_excel(w, sheet_name="Overview", index=False)
        df.isnull().sum().reset_index().rename(columns={"index":"Column",0:"Missing"}).to_excel(w, sheet_name="Missing", index=False)
        df.describe(include="all").transpose().to_excel(w, sheet_name="Stats")

        wb = w.book
        s = wb.add_worksheet("Charts")
        w.sheets["Charts"] = s

        num = df.select_dtypes(include="number").columns
        if len(num) >= 2:
            plt.figure(figsize=(6,5))
            sns.heatmap(df[num].corr(), annot=True, cmap="coolwarm", fmt=".2f")
            plt.tight_layout()
            path = f"{charts}/heat.png"
            plt.savefig(path)
            plt.close()
            s.insert_image("A1", path)
        else:
            s.write("A1", "No numeric data")

        cat = df.select_dtypes(include=["object","category"]).columns[:3]
        y = 25
        for c in cat:
            plt.figure(figsize=(6,4))
            df[c].value_counts().head(10).plot(kind="bar", title=c)
            plt.tight_layout()
            p = f"{charts}/{c}.png"
            plt.savefig(p)
            plt.close()
            s.insert_image(f"A{y}", p)
            y += 20

    shutil.rmtree(charts, ignore_errors=True)
    with open(name, "rb") as f:
        return f.read()

if f:
    df = pd.read_csv(f)
    df = clean_df(df)
    st.write("Rows:", len(df), "Cols:", len(df.columns))
    st.dataframe(df.head())

    st.subheader("Stats")
    try:
        st.dataframe(df.describe(include='all').transpose())
    except:
        st.warning("Could not calculate some stats")

    num = df.select_dtypes(include="number").columns
    if len(num) >= 2:
        st.subheader("Correlation")
        fig, ax = plt.subplots(figsize=(6,5))
        sns.heatmap(df[num].corr(), annot=True, cmap="coolwarm", fmt=".2f", ax=ax)
        st.pyplot(fig)

    data = make_report(df)
    st.download_button("Download Excel Report", data, "report.xlsx")
else:
    st.info("Upload a CSV to start")
