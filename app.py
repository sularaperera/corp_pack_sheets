import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Warehouse Sheet Generator",
    layout="wide"
)

st.title("📦 Warehouse Sheet Generator for Corporate Orders")

st.markdown("""
Upload the Mintsoft and DSD files to generate:

- Label Sheet
- Milk Sheet
- Pick Sheets
""")

# -----------------------------
# File Uploads
# -----------------------------
mintsoft_file = st.file_uploader(
    "Upload Mintsoft File",
    type=["xlsx"]
)

dsd_file = st.file_uploader(
    "Upload DSD File",
    type=["xlsx"]
)

# -----------------------------
# Helper Function
# -----------------------------
def to_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    output.seek(0)
    return output

# -----------------------------
# Process Files
# -----------------------------
if mintsoft_file and dsd_file:

    try:
        # Read files
        mintsoft_df = pd.read_excel(mintsoft_file)
        dsd_df = pd.read_excel(dsd_file)

        # -----------------------------
        # Mintsoft cleanup
        # -----------------------------
        mintsoft_df = mintsoft_df[
            [
                "Product SKU",
                "Product Name",
                "Order Number",
                "FirstName",
                "Address 1",
                "Courier Service",
                "Product Qty"
            ]
        ]

        # -----------------------------
        # DSD cleanup
        # -----------------------------
        dsd_df = dsd_df[
            [
                "Address Line One",
                "Sequence"
            ]
        ]

        # -----------------------------
        # Clean Address Fields
        # -----------------------------
        mintsoft_df["Address 1 Clean"] = (
            mintsoft_df["Address 1"]
            .astype(str)
            .str.strip()
            .str.lower()
        )

        dsd_df["Address Line One Clean"] = (
            dsd_df["Address Line One"]
            .astype(str)
            .str.strip()
            .str.lower()
        )

        # -----------------------------
        # Merge
        # -----------------------------
        final_df = mintsoft_df.merge(
            dsd_df,
            left_on="Address 1 Clean",
            right_on="Address Line One Clean",
            how="left"
        )

        # Remove helper columns
        final_df = final_df.drop(
            columns=[
                "Address 1 Clean",
                "Address Line One",
                "Address Line One Clean"
            ]
        )

        # -----------------------------
        # Pick Sheet
        # -----------------------------
        pick_df = final_df.sort_values(
            by=[
                "Product Name",
                "Courier Service",
                "Sequence"
                ]
        )

        # Rearrange columns
        pick_df = pick_df[
            [
                "Product SKU",
                "Product Name",
                "Product Qty",
                "Order Number",
                "FirstName",
                "Address 1",
                "Courier Service",
                "Sequence"
            ]
        ]

        # -----------------------------
        # Label Sheet
        # -----------------------------
        labels_df = final_df[
            [
                "Order Number",
                "FirstName",
                "Address 1",
                "Courier Service",
                "Sequence"
            ]
        ]

        labels_df = labels_df.drop_duplicates(
            subset=["Order Number"]
        )

        # -----------------------------
        # Milk Sheet
        # -----------------------------
        milk_products_df = final_df [
            final_df["Product Name"]
            .astype(str)
            .str.contains(r"milk|meadowfresh", case=False, na=False)
        ]

        milk_df = milk_products_df.pivot_table(
            index=["Product SKU", "Product Name"],
            columns="Courier Service",
            values="Product Qty",
            aggfunc="sum",
            fill_value=0
        ).reset_index()

        st.success("Files processed successfully!")

        # -----------------------------
        # Tabs
        # -----------------------------
        tab1, tab2, tab3 = st.tabs([
            "📋 Label Sheet",
            "🥛 Milk Sheet",
            "📦 Pick Sheets"
        ])

        # -----------------------------
        # Label Sheet Tab
        # -----------------------------
        with tab1:
            st.dataframe(labels_df, use_container_width=True)

            st.download_button(
                label="⬇ Download Label Sheet",
                data=to_excel(labels_df),
                file_name="Label Sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # -----------------------------
        # Milk Sheet Tab
        # -----------------------------
        with tab2:
            st.dataframe(milk_df, use_container_width=True)

            st.download_button(
                label="⬇ Download Milk Sheet",
                data=to_excel(milk_df),
                file_name="Milk Sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # -----------------------------
        # Pick Sheets Tab
        # -----------------------------
        with tab3:
            st.dataframe(pick_df, use_container_width=True)

            st.download_button(
                label="⬇ Download Pick Sheets",
                data=to_excel(pick_df),
                file_name="Pick Sheets.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error processing files: {e}")

else:
    st.info("Please upload both files to continue.")