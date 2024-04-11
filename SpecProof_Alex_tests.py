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
    ################## DSBP DATA PREPARATION ####################
    ## IMPORT EXCEL FILES
    # filename_dsbp = "Digital_Base_Plan_Test.xlsx"
    df_pipo_fpc = pd.read_excel(filename_dsbp, sheet_name="PIPO FPC Data")
    df_full_bom = pd.read_excel(filename_dsbp, sheet_name="Full BOM - Materials")
    df_details_mat = pd.read_excel(filename_dsbp, sheet_name="Packaging Details - Materials")
    df_stratcount = pd.read_excel(filename_dsbp, sheet_name="Strategies & Counts")

    ## DATASETS FILTERING
    dsbp_pipo_fpc_keep = ["PI FPC Code (IOL)", "Consumer Unit Size: As Sourced (IL/PM)", "Consumer Unit Size: As Sourced UoM (IL/PM)", "PI IT/CS Count Value (IL/PM)", "Primary Production Plant (One) (IL/PM)", "Category (IL/PM)", "Primary Product Package Type (IL/PM)", "Brand (IL/PM)"]
    dsbp_full_bom_keep = ["PI FPC Code (IL/PM)", "PI FPC Description (IL/PM)", "PI Material Number (IL/PM)", "Material Description (IL/PM)", "Material Type (TRL)", "Material Quantity (Package SPOC)", "Substitution Material? (Package SPOC)", "Substitution Primary Material (TRL)", "Requires AW (IL/PM)", "PI Plant Code(s) (IL/PM)", "Markets (IL/PM)"]
    dsbp_details_mat_keep = ["PI Packaging Component Type (TRL)", "PI Material GCAS (TRL)", "PI Material Description (TRL)", "Material type (IL/PM)", "Master (IL/PM)", "Primary TD (ILST) (IL/PM)", "Requires AW (TRL)", "MEP name (TRL)", "MEP description (TRL)"]
    dsbp_stratcount_keep = ["PI FPC Code (IL/PM, IOL)", "Production Plant (Primary) Pallet (IL/PM, IOL)"]
    df_pipo_fpc = df_pipo_fpc.drop(df_pipo_fpc.columns.difference(dsbp_pipo_fpc_keep), axis=1)
    df_full_bom = df_full_bom.drop(df_full_bom.columns.difference(dsbp_full_bom_keep), axis=1)
    df_details_mat = df_details_mat.drop(df_details_mat.columns.difference(dsbp_details_mat_keep), axis = 1)
    df_stratcount = df_stratcount.drop(df_stratcount.columns.difference(dsbp_stratcount_keep), axis=1)

    ## DATASETS MERGE
    df_full_bom = df_full_bom.merge(df_details_mat, left_on="PI Material Number (IL/PM)", right_on="PI Material GCAS (TRL)")
    # df_full_bom.to_excel("specproof_output2.xlsx", sheet_name="Full BOM")
    df_full_bom = df_full_bom.merge(df_pipo_fpc, left_on="PI FPC Code (IL/PM)", right_on="PI FPC Code (IOL)")
    df_full_bom = df_full_bom.merge(df_stratcount, left_on="PI FPC Code (IL/PM)", right_on="PI FPC Code (IL/PM, IOL)")
    df_full_bom = df_full_bom.dropna(axis=0, subset=["PI Material Number (IL/PM)", "PI FPC Code (IOL)"])
    df_full_bom = df_full_bom.drop_duplicates()
    df_full_bom["Marketing Size"] = df_full_bom["Consumer Unit Size: As Sourced (IL/PM)"].astype("str") + df_full_bom["Consumer Unit Size: As Sourced UoM (IL/PM)"]
    df_full_bom = df_full_bom.drop(["Consumer Unit Size: As Sourced (IL/PM)", "Consumer Unit Size: As Sourced UoM (IL/PM)", "PI FPC Code (IL/PM, IOL)", "PI FPC Code (IOL)", "PI Material GCAS (TRL)", "PI Material Description (TRL)", "Material type (IL/PM)", "Requires AW (IL/PM)"], axis=1)
    df_full_bom = df_full_bom.reset_index()
    df_full_bom = df_full_bom.rename(columns={"PI FPC Code (IL/PM)": "FPP Name", "PI FPC Description (IL/PM)": "FPP Title", "PI Material Number (IL/PM)": "Material Number", "Material Description (IL/PM)": "Material Title", "Material Type (TRL)": "Material Type", "Material Quantity (Package SPOC)": "Material Quantity", "Substitution Material? (Package SPOC)": "Material Substitution?", "Substitution Primary Material (TRL)": "Substitution Number", "PI Plant Code(s) (IL/PM)": "Producing Plant Code", "Markets (IL/PM)": "Market of Sale(s)", "PI Packaging Component Type (TRL)": "Pack Component Type", "Master (IL/PM)": "Master Name", "Primary TD (ILST) (IL/PM)": "TD/ILST", "Requires AW (TRL)": "Has Art?", "MEP name (TRL)": "Supplier/MEP Name", "MEP description (TRL)": "Supplier/MEP Title", "Category (IL/PM)": "Category", "Brand (IL/PM)": "Brand", "Primary Product Package Type (IL/PM)": "Primary Pack Type", "Consumer Unit Size: As Sourced (IL/PM)": "Marketing Size", "Consumer Unit Size: As Sourced UoM (IL/PM)": "Size UoM", "PI IT/CS Count Value (IL/PM)": "Case Count", "Primary Production Plant (One) (IL/PM)": "Producing Plant Name", "Production Plant (Primary) Pallet (IL/PM, IOL)": "Pallet Type"})
    df_full_bom["Material Type"] = df_full_bom["Material Type"].replace(["PMP", "FOP"], ["Packaging Material Part", "Formulation Part"])
    df_full_bom["Material Substitution?"] = df_full_bom["Material Substitution?"].fillna("No")

    ## EXPORT TO EXCEL
    # df_full_bom.to_excel("specproof_output.xlsx", sheet_name="Full BOM")
    ## Display the df_full_bom DataFrame on Streamlit
    st.write("DSBP BOM Data:")
    st.dataframe(df_full_bom)



    ################## DSM DATA PREPARATION ####################
    ## FUNCTIONS
    def new_column_based_on_list(df1, new_column_name, list):
        df2 = df1
        df2[new_column_name] = ""
        for i in range(len(df1.index)):
            for j in range(len(list)):
                if df2.iloc[i]["Name/Number"] == list[j][0]:
                    df2.at[i, new_column_name] = list[j][1]
        return df2
        
    ## IMPORT EXCEL FILES
    # filename_dsm = "DSM_Test.xlsx"
    df_dsm_attributes = pd.read_excel(filename_dsm, sheet_name="Attributes")
    df_dsm_bom = pd.read_excel(filename_dsm, sheet_name="Bill of Materials")
    df_dsm_ebom = pd.read_excel(filename_dsm, sheet_name="EBOM W&D")
    df_dsm_plant = pd.read_excel(filename_dsm, "Plants")
    df_dsm_wd = pd.read_excel(filename_dsm, sheet_name="Weights & Dimensions")
    df_dsm_mos = pd.read_excel(filename_dsm, sheet_name="Market Of Sale")
    df_dsm_mep = pd.read_excel(filename_dsm, sheet_name="Part Component Equivalents")
    df_dsm_sd = pd.read_excel(filename_dsm, sheet_name="Specs and Docs")

    ## PROCESSING
    fpp_list = df_dsm_bom["Name/Number"].unique()
    case_count_list = []
    for x in fpp_list:
        df_fpp_bom = df_dsm_bom[df_dsm_bom["Name/Number"] == x]
        df_cop_bom = df_fpp_bom[df_fpp_bom["Material Type"] == "Consumer Unit Part"]
        df_cop_bom = df_cop_bom.reset_index()
        case_count = df_cop_bom.iloc[0]["Qty"]
        case_count_list.append([x,case_count])

    df_dsm_bom = new_column_based_on_list(df_dsm_bom, "Case count/COP Qty", case_count_list)

    df_dsm_plant[["Primary Plant", "Plant Code"]] = df_dsm_plant.Plants.str.split("~", expand = True)

    df_dsm_wd = df_dsm_wd.dropna(axis=0, subset=["Transport Unit - Name"])
    df_dsm_wd = df_dsm_wd[df_dsm_wd["TransportUnit - Include In SAP BOM Feed"] == "Yes"]  # For now only considers the main pallet type, to add new additional pallet types, I would concatenate inside of 1 cell in descending

    mos_combined = df_dsm_mos.groupby(['Name/Number']).apply(lambda group: ', '.join(group['Market of Sale']))
    df_dsm_mos = pd.DataFrame(mos_combined)
    df_dsm_mos = df_dsm_mos.reset_index()
    df_dsm_mos = df_dsm_mos.rename(columns={0: "Market of Sale"})

    df_dsm_bom = df_dsm_bom[df_dsm_bom["Type"] == "Finished Product Part"].astype(({"Name/Number": "str"}))
    df_dsm_ebom = df_dsm_ebom[df_dsm_ebom["Name/Number"].str.contains("CUP|COP") == False].astype(({"Name/Number": "str"}))
    df_dsm_plant = df_dsm_plant[df_dsm_plant["Type"] == "Finished Product Part"].astype(({"Name/Number": "int32"}))
    df_attributes_fpp = df_dsm_attributes[df_dsm_attributes["Type"] == "Finished Product Part"].astype(({"Name/Number": "str"}))
    df_attributes_pmp = df_dsm_attributes[df_dsm_attributes["Type"].isin(["Packaging Material Part", "Formulation Part"])]

    df_dsm_mep = df_dsm_mep.dropna(axis=0, subset=["MEP Name"])

    df_dsm_sd = df_dsm_sd[df_dsm_sd["Master Specification-Type"] == "Illustration"]
    sd_combined = df_dsm_sd.groupby(['Name/Number']).apply(lambda group: ', '.join(group['Master Specification-Name']))
    df_dsm_sd = pd.DataFrame(sd_combined)
    df_dsm_sd = df_dsm_sd.reset_index()
    df_dsm_sd = df_dsm_sd.rename(columns={0: "Master Specification-Name"})

    ## DATASETS FILTERING
    dsm_att_fpp_keep = ["Name/Number", "Segment", "Brand", "Primary Packaging Type", "Primary Organization", "Marketing Size"]
    dsm_att_pmp_keep = ["Name/Number", "Class", "Reported Function", "Printing Process", "Gross Weight", "Weight UoM", "Packaging Component Type"]
    dsm_bom_keep = ["Name/Number", "Title", "Material Type", "Material Number", "Material Title", "Qty", "Substitution", "Substitute Name/Number", "Has Art", "Case count/COP Qty"]
    dsm_ebom_keep = ["Name/Number", "Material Number", "Master Part Name", "Master Title"]
    dsm_plant_keep = ["Name/Number", "Primary Plant", "Plant Code"]
    dsm_wd_keep = ["Name/Number", "Transport Unit - Name", "Transport Unit - Pallet Type", "Transport Unit - Stacking Pattern GCAS Code"]
    dsm_mos_keep = ["Name/Number", "Market of Sale"]
    dsm_mep_keep = ["Name/Number", "MEP Name", "MEP Manufacturer", "MEP Artwork Primary"]
    dsm_sd_keep = ["Name/Number", "Master Specification-Name"]
    df_dsm_bom = df_dsm_bom.drop(df_dsm_bom.columns.difference(dsm_bom_keep), axis=1)
    df_dsm_ebom = df_dsm_ebom.drop(df_dsm_ebom.columns.difference(dsm_ebom_keep), axis=1)
    df_dsm_plant = df_dsm_plant.drop(df_dsm_plant.columns.difference(dsm_plant_keep), axis=1).astype(({"Name/Number": "str"}))
    df_dsm_wd = df_dsm_wd.drop(df_dsm_wd.columns.difference(dsm_wd_keep), axis=1).astype(({"Name/Number": "str"}))
    df_dsm_mos = df_dsm_mos.drop(df_dsm_mos.columns.difference(dsm_mos_keep), axis=1).astype(({"Name/Number": "str"}))
    df_attributes_fpp = df_attributes_fpp.drop(df_attributes_fpp.columns.difference(dsm_att_fpp_keep), axis=1)
    df_attributes_pmp = df_attributes_pmp.drop(df_attributes_pmp.columns.difference(dsm_att_pmp_keep), axis=1)
    df_dsm_mep = df_dsm_mep.drop(df_dsm_mep.columns.difference(dsm_mep_keep), axis=1).astype(({"Name/Number": "str"}))
    df_dsm_sd = df_dsm_sd.drop(df_dsm_sd.columns.difference(dsm_sd_keep), axis=1).astype(({"Name/Number": "str"}))

    ## DATASETS MERGE
    df_dsm_bom = df_dsm_bom.merge(df_dsm_ebom, left_on=["Name/Number", "Material Number"], right_on=["Name/Number", "Material Number"]).astype(({"Name/Number": "str"}))
    df_dsm_bom = df_dsm_bom.merge(df_dsm_plant, left_on=["Name/Number"], right_on=["Name/Number"]).astype(({"Name/Number": "str"}))
    df_dsm_bom = df_dsm_bom.merge(df_dsm_wd, left_on=["Name/Number"], right_on=["Name/Number"]).astype(({"Name/Number": "str"}))
    df_dsm_bom = df_dsm_bom.merge(df_dsm_mos, left_on=["Name/Number"], right_on=["Name/Number"]).astype(({"Name/Number": "str"}))
    df_dsm_bom = df_dsm_bom.merge(df_attributes_fpp, left_on=["Name/Number"], right_on=["Name/Number"]).astype(({"Name/Number": "str"}))
    df_dsm_bom = df_dsm_bom.dropna(axis=0, subset=["Material Number"])
    df_dsm_bom = df_dsm_bom.merge(df_attributes_pmp, how="left", left_on=["Material Number"], right_on=["Name/Number"], suffixes=(None, "_overlap"))
    df_dsm_bom = df_dsm_bom.merge(df_dsm_mep, how="left", left_on=["Material Number"], right_on=["Name/Number"], suffixes=(None, "_overlap2"))
    df_dsm_bom = df_dsm_bom.merge(df_dsm_sd, how="left", left_on=["Material Number"], right_on=["Name/Number"], suffixes=(None, "_overlap3"))
    df_dsm_bom = df_dsm_bom.drop(["Name/Number_overlap", "Name/Number_overlap2", "Name/Number_overlap3"], axis=1)
    df_dsm_bom = df_dsm_bom.rename(columns={"Name/Number": "FPP Name", "Title": "FPP Title", "Material Type": "Material Type", "Material Number": "Material Number", "Material Title": "Material Title", "Qty": "Material Quantity", "Substitution": "Material Substitution?", "Substitute Name/Number": "Substitution Number", "Has Art": "Has Art?", "Case count/COP Qty": "Case Count", "Master Part Name": "Master Name", "Master Title": "Master Title", "Primary Plant": "Producing Plant Name", "Plant Code": "Producing Plant Code", "Transport Unit - Name": "TUP Name", "Transport Unit - Pallet Type": "Pallet Type", "Transport Unit - Stacking Pattern GCAS Code": "SPS Name", "Market of Sale": "Market of Sale(s)", "Marketing Size": "Marketing Size", "Segment": "Category", "Brand": "Brand", "Primary Packaging Type": "Primary Pack Type", "Primary Organization": "Primary Organization", "Class": "Class", "Reported Function": "Reported Function", "Printing Process": "Printing Process", "Gross Weight": "Pack Gross Weight", "Weight UoM": "Weight UoM", "Packaging Component Type": "Pack Component Type", "MEP Name": "Supplier/MEP Name", "MEP Manufacturer": "Supplier/MEP Title", "Master Specification-Name": "TD/ILST"})
    df_dsm_bom["Category"] = df_dsm_bom["Category"].map(lambda x: x.lstrip("All Other"))
    df_dsm_bom["Has Art?"] = df_dsm_bom["Has Art?"].replace([True, False], ["Yes", "No"])
    df_dsm_bom["Has Art?"] = df_dsm_bom["Has Art?"].fillna("No")
    df_dsm_bom = df_dsm_bom.drop(df_dsm_bom[(df_dsm_bom["Has Art?"] == "Yes") & (df_dsm_bom["MEP Artwork Primary"] == "No")].index)
    df_dsm_bom = df_dsm_bom.drop_duplicates()
    df_dsm_bom = df_dsm_bom.reset_index()

    ## EXPORT TO EXCEL
    # df_dsm_bom.to_excel("specproof_dsm_output.xlsx", sheet_name="DSM BOM")
    ## Display the df_full_bom DataFrame on Streamlit
    st.write("DSM BOM data:")
    st.dataframe(df_dsm_bom)


    ################## COMBINE & COMPARE ####################
    # FUNCTIONS
    def highlight_diff(data, color='yellow'):
        attr = 'background-color: {}'.format(color)
        other1 = data.xs('DSBP', axis='columns', level=-1)
        other2 = data.xs('DSM', axis='columns', level=-1)
        return pd.DataFrame(np.where(data.ne(other1, level=0) | data.ne(other2, level=0), attr,''), index=data.index, columns=data.columns)

    df_dsm_stripped = df_dsm_bom[df_dsm_bom["Material Type"].isin(["Packaging Material Part", "Formulation Part"])]
    df_dsm_stripped = df_dsm_stripped.astype(({"FPP Name": "int32", "Material Number": "int32"}))
    df_dsm_stripped = df_dsm_stripped.sort_values(["FPP Name", "Material Number"], ascending=True)
    df_dsm_stripped = df_dsm_stripped.reset_index()
    df_full_bom = df_full_bom.astype(({"FPP Name": "int32", "Material Number": "int32"}))
    df_full_bom = df_full_bom.sort_values(["FPP Name", "Material Number"], ascending=True)
    df_full_bom = df_full_bom.reset_index()
    df_compare = df_full_bom.merge(df_dsm_stripped, how="outer", left_on=["FPP Name", "Material Number"], right_on=["FPP Name", "Material Number"], suffixes=("_DSBP", "_DSM"))
    #df_compare = pd.concat([df_full_bom, df_dsm_stripped], axis=1, keys=["DSBP", "DSM"])
    df_compare = pd.concat([df_compare[['FPP Name','FPP Title_DSBP','FPP Title_DSM','Material Number','Material Title_DSBP','Material Title_DSM','Material Type_DSBP','Material Type_DSM']], df_compare[df_compare.columns.difference(['FPP Name','FPP Title_DSBP','FPP Title_DSM','Material Number','Material Title_DSBP','Material Title_DSM','Material Type_DSBP','Material Type_DSM'])].sort_index(axis=1)],ignore_index=False,axis=1)
    
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
    df_compare_styled = df_compare.style.apply(highlight_diff, axis=1)
    
    print(df_compare.head())
    # df_compare.to_excel("specproof_compare_output.xlsx", sheet_name="Comparison")
    # Display the df_compare DataFrame on Streamlit
    st.write("ERAT Comparison:")
    st.dataframe(df_compare_styled)



def SpecFlow():
    st.title("SpecFlow")
    code = '''print("Still in construction...")'''
    st.code(code, language='python')
    st.header("", anchor=None, help=None, divider="grey")



def SpecCatalog():
    st.title("SpecCatalog")
    st.write("For now, you can access the catalog via:")
    st.write("https://app.powerbi.com/groups/6ac2bd6a-88a6-4dd5-b551-3073c566a784/reports/0838e724-65e2-4c89-87ea-df227700c95c/ReportSection4bd0a5cf1d2a65ed2d24?experience=power-bi")
    st.header("", anchor=None, help=None, divider="grey")
    st.markdown(
    """
    <iframe src="https://app.powerbi.com/groups/6ac2bd6a-88a6-4dd5-b551-3073c566a784/reports/0838e724-65e2-4c89-87ea-df227700c95c/ReportSection4bd0a5cf1d2a65ed2d24?experience=power-bi" width="100%" height="800"></iframe>
    """,
    unsafe_allow_html=True
    )




# Main function
def main():
    with st.sidebar:
        selected = option_menu(
            menu_title="Navigation",
            options=["ERAT", "SpecFlow", "SpecCatalog"],
            icons=["1-circle-fill", "2-circle-fill", "bar-chart-line-fill"],
            menu_icon="list",
            default_index=0,
        )

    # Render the selected page
    if selected == "ERAT":
        ERAT()
    elif selected == "SpecFlow":
        SpecFlow()
    elif selected == "SpecCatalog":
        SpecCatalog()

if __name__ == "__main__":
    main()