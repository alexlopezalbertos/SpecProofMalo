## LIBRARY IMPORTS
import pandas as pd
import numpy as np
import streamlit as st
from streamlit_option_menu import option_menu

def ERAT():
    st.title("ERAT")
    # UPLOAD BOTH EXCEL FILES #
    filename_dsbp = st.file_uploader("Upload DSBP report")
    filename_dsm = st.file_uploader("Upload ENOVIA DSM report")

    if st.button("Run ERAT", type="primary"):
        if filename_dsbp is None or filename_dsm is None:
            # Identify which inputs are missing
            missing_inputs = []
            if filename_dsbp is None:
                missing_inputs.append(":red[DSBP report]")
            if filename_dsm is None:
                missing_inputs.append(":red[ENOVIA DSM report]")
            # Display a warning message
            st.warning(f"Please upload: {', '.join(missing_inputs)}.", icon="⚠️")
        else:
            run_ERAT(filename_dsbp, filename_dsm)

def run_ERAT(filename_dsbp, filename_dsm):
    # READ DSBP BOM #
    df_dsbp_bom = pd.read_excel(filename_dsbp, sheet_name="Full BOM - Materials")
    df_dsbp_bom_keep = ["PI FPC Code (IL/PM)", "PI FPC Description (IL/PM)", "PI Material Number (IL/PM)", "Material Description (IL/PM)", "Material Type (TRL)"]
    df_dsbp_bom = df_dsbp_bom.drop(df_dsbp_bom.columns.difference(df_dsbp_bom_keep), axis=1)
    df_dsbp_bom = df_dsbp_bom.rename(columns={"PI FPC Code (IL/PM)": "FPP Name", "PI FPC Description (IL/PM)": "FPP Title_DSBP", "PI Material Number (IL/PM)": "Material Number",
                                               "Material Description (IL/PM)": "Material Title_DSBP", "Material Type (TRL)": "Material Type_DSBP"})
    df_dsbp_bom["Material Type_DSBP"] = df_dsbp_bom["Material Type_DSBP"].replace(["PMP", "FOP", "RMP"], ["Packaging Material Part", "Formulation Part", "Raw Material Part"])
    st.dataframe(df_dsbp_bom)

    # READ DSM BOM #
    df_dsm_bom = pd.read_excel(filename_dsm, sheet_name="Bill of Materials")
    df_dsm_bom = df_dsm_bom[df_dsm_bom["Type"] == "Finished Product Part"]
    df_dsm_bom = df_dsm_bom[df_dsm_bom["Material Type"].isin(["Packaging Material Part", "Formulation Part", "Raw Material Part"])]
    df_dsm_bom_keep = ["Name/Number", "Title", "Material Number", "Material Title", "Material Type"]
    df_dsm_bom = df_dsm_bom.drop(df_dsm_bom.columns.difference(df_dsm_bom_keep), axis=1)
    df_dsm_bom = df_dsm_bom.reset_index()
    df_dsm_bom = df_dsm_bom.rename(columns={"Name/Number": "FPP Name", "Title": "FPP Title_DSM", "Material Number": "Material Number",
                                             "Material Title": "Material Title_DSM", "Material Type": "Material Type_DSM"})
    st.dataframe(df_dsm_bom)

    df_dsm_bom = df_dsm_bom.astype(str)
    df_dsbp_bom = df_dsbp_bom.astype(str)

    df_merged_bom = pd.merge(df_dsbp_bom, df_dsm_bom, on=['FPP Name', 'Material Number'], how='outer')
    column_order = ['FPP Name', 'FPP Title_DSBP', 'FPP Title_DSM', 'Material Number', 'Material Title_DSBP', 'Material Title_DSM', 'Material Type_DSBP', 'Material Type_DSM']
    df_merged_bom = df_merged_bom[column_order]

    st.dataframe(df_merged_bom)



    def highlight_diff(row):
        color_map = []
        for col in row.index:
            if col.endswith("_DSBP") or col.endswith("_DSM"):
                col_dsbp = col.replace("_DSM", "_DSBP")
                col_dsm = col.replace("_DSBP", "_DSM")
                if row[col_dsbp] != row[col_dsm]:
                    color_map.append("background-color: lightcoral")
                else:
                    color_map.append("background-color: lightgreen")
            else:
                color_map.append("")
        return color_map

    # Apply the highlight_diff function to each row of the DataFrame
    df_compare_styled = df_merged_bom.style.apply(highlight_diff, axis=1)
    
    print(df_merged_bom.head())
    # df_compare.to_excel("specproof_compare_output.xlsx", sheet_name="Comparison")
    # Display the df_compare DataFrame on Streamlit
    st.write("ERAT Comparison:")
    st.dataframe(df_compare_styled)



# Main function
def main():
    with st.sidebar:
        selected = option_menu(
            menu_title="Navigation",
            options=["ERAT"],
            icons=["1-circle-fill"],
            menu_icon="list",
            default_index=0,
        )

    # Render the selected page
    if selected == "ERAT":
        ERAT()

if __name__ == "__main__":
    main()