import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="DataFix",
    page_icon="🛠️",
    layout="wide"
)

st.title("🛠️ DataFix")
st.subheader("Autonomous Data Quality and Repair Engine")
st.write("Detect, diagnose and repair your datasets.")

uploaded_file = st.file_uploader(
    "Upload your CSV dataset",
    type=["csv"]
)

if uploaded_file is not None:
    try:
        df = pd.read_csv(uploaded_file)

        st.success("Dataset uploaded successfully!")

        st.header("Dataset Overview")

        col1, col2, col3 = st.columns(3)
        col1.metric("Total Rows", len(df))
        col2.metric("Total Columns", len(df.columns))
        col3.metric("Total Cells", df.size)

        st.subheader("Dataset Preview")
        st.dataframe(df)

        missing = int(df.isnull().sum().sum())
        duplicates = int(df.duplicated().sum())

        st.header("Data Quality Problems")

        col1, col2 = st.columns(2)
        col1.metric("Missing Values", missing)
        col2.metric("Duplicate Rows", duplicates)

        st.subheader("Missing Values by Column")
        st.dataframe(df.isnull().sum().reset_index())

        if st.button("Repair Data"):
            cleaned = df.copy()

            for column in cleaned.columns:
                if cleaned[column].isnull().sum() > 0:
                    if pd.api.types.is_numeric_dtype(cleaned[column]):
                        cleaned[column] = cleaned[column].fillna(
                            cleaned[column].median()
                        )
                    else:
                        cleaned[column] = cleaned[column].fillna("Unknown")

            cleaned = cleaned.drop_duplicates().reset_index(drop=True)

            st.session_state["cleaned"] = cleaned
            st.success("Data repair completed!")

        if "cleaned" in st.session_state:
            cleaned = st.session_state["cleaned"]

            st.header("Repair Results")

            st.metric(
                "Remaining Missing Values",
                int(cleaned.isnull().sum().sum())
            )

            st.metric(
                "Remaining Duplicate Rows",
                int(cleaned.duplicated().sum())
            )

            st.subheader("Cleaned Dataset")
            st.dataframe(cleaned)

            csv = cleaned.to_csv(index=False).encode("utf-8")

            st.download_button(
                "Download Cleaned CSV",
                csv,
                "datafix_cleaned.csv",
                "text/csv"
            )

    except Exception as e:
        st.error(f"Error: {e}")