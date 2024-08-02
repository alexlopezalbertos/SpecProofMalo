import pandas as pd
import numpy as np
import streamlit as st
from streamlit_option_menu import option_menu

# Define the correct password
CORRECT_PASSWORD = "ESSMPD"

def SPECPROOF():
    st.title("SpecProof")
    st.write("")

    # Check if the password is already verified
    if "password_verified" not in st.session_state:
        st.session_state.password_verified = False

    # If the password is not verified, show the password input
    if not st.session_state.password_verified:
        password = st.text_input(":lock: Insert password", type='password', help="Request password to @Lopez, Alejandro")
        if password == CORRECT_PASSWORD:
            st.session_state.password_verified = True
            st.rerun()  # Rerun the app to show the rest of the content
        else:
            if password:
                st.error("Incorrect password")
    else:
        # UPLOAD BOTH EXCEL FILES #
        st.subheader(":page_facing_up: DSBP", divider="blue")
        filename_dsbp = st.file_uploader("Upload DSBP report")
        st.write("")
        st.subheader(":page_facing_up: ENOVIA", divider="blue")
        filename_dsm = st.file_uploader("Upload DSM report")
        st.write("")
        st.subheader(":page_facing_up: TRACKSHEET", divider="blue")
        st.markdown(":ledger: Optional step, only upload the Tracksheet if you want to compare it against Enovia's TUP & SPS")
        filename_tracksheet = st.file_uploader("Upload Tracksheet")
        st.write("")

        if st.button("RUN", type="primary"):
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
                run_SPECPROOF(filename_dsbp, filename_dsm, filename_tracksheet)

def run_SPECPROOF(filename_dsbp, filename_dsm, filename_tracksheet):
   
    # ========================================================================== BOM ==================================================================================================
    
    st.subheader(":one: BOM:", divider="blue")

    # READ DSBP BOM #
    df_dsbp_bom = pd.read_excel(filename_dsbp, sheet_name="Full BOM - Materials")
    df_dsbp_bom_keep = ["PI FPC Code (IL/PM)", "PI FPC Description (IL/PM)", "PI Material Number (IL/PM)", "Material Description (IL/PM)", "Material Type (TRL)"]
    df_dsbp_bom = df_dsbp_bom.drop(df_dsbp_bom.columns.difference(df_dsbp_bom_keep), axis=1)
    df_dsbp_bom = df_dsbp_bom.rename(columns={"PI FPC Code (IL/PM)": "FPP Name", "PI FPC Description (IL/PM)": "FPP Title_DSBP", "PI Material Number (IL/PM)": "Material Number",
                                               "Material Description (IL/PM)": "Material Title_DSBP", "Material Type (TRL)": "Material Type_DSBP"})
    df_dsbp_bom = df_dsbp_bom[df_dsbp_bom["Material Type_DSBP"].isin(["PMP", "FOP", "RMP", "APP"])]
    df_dsbp_bom["Material Type_DSBP"] = df_dsbp_bom["Material Type_DSBP"].replace(["PMP", "FOP", "RMP", "APP"], ["Packaging Material Part", "Formulation Part", "Raw Material Part", "Assembled Product Part"])
    df_dsbp_bom["Material Number"] = pd.to_numeric(df_dsbp_bom["Material Number"], errors='coerce').fillna(0).astype(int)
    df_dsbp_bom = df_dsbp_bom.astype(str)
    # st.dataframe(df_dsbp_bom)

    # READ DSM BOM #
    df_dsm_bom = pd.read_excel(filename_dsm, sheet_name="Bill of Materials")
    df_dsm_bom = df_dsm_bom[df_dsm_bom["Type"] == "Finished Product Part"]
    df_dsm_bom = df_dsm_bom[df_dsm_bom["Material Type"].isin(["Packaging Material Part", "Formulation Part", "Raw Material Part", "Assembled Product Part", "Raw Material"])]
    df_dsm_bom_keep = ["Name/Number", "Title", "Material Number", "Material Title", "Material Type"]
    df_dsm_bom = df_dsm_bom.drop(df_dsm_bom.columns.difference(df_dsm_bom_keep), axis=1)
    df_dsm_bom = df_dsm_bom.reset_index()
    df_dsm_bom = df_dsm_bom.rename(columns={"Name/Number": "FPP Name", "Title": "FPP Title_DSM", "Material Number": "Material Number",
                                             "Material Title": "Material Title_DSM", "Material Type": "Material Type_DSM"})
    df_dsm_bom["Material Type_DSM"] = df_dsm_bom["Material Type_DSM"].replace(["Raw Material"], ["Raw Material Part"])
    df_dsm_bom = df_dsm_bom.astype(str)
    # st.dataframe(df_dsm_bom)  
    
    # MERGE BOM #
    df_merged_bom = pd.merge(df_dsbp_bom, df_dsm_bom, on=['FPP Name', 'Material Number'], how='outer')
    column_order = ['FPP Name', 'FPP Title_DSBP', 'FPP Title_DSM', 'Material Number', 'Material Title_DSBP', 'Material Title_DSM', 'Material Type_DSBP', 'Material Type_DSM']
    df_merged_bom = df_merged_bom[column_order]
    df_merged_bom["Material Title_DSBP"] = df_merged_bom["Material Title_DSBP"].str.strip() #remove last space after title
    df_merged_bom["Material Title_DSM"] = df_merged_bom["Material Title_DSM"].str.strip() #remove last space after title
    df_merged_bom["Material Title_DSBP"] = df_merged_bom["Material Title_DSBP"].str.replace('\u00A0', ' ') #Replaces weird spaces for normal spaces
    df_merged_bom["Material Title_DSM"] = df_merged_bom["Material Title_DSM"].str.replace('\u00A0', ' ') #Replaces weird spaces for normal spaces
    df_merged_bom = df_merged_bom.sort_values(by='FPP Name', ascending=True)
    df_merged_bom = df_merged_bom.reset_index()
    df_merged_bom.drop("index", axis=1, inplace=True)
    # st.dataframe(df_merged_bom)

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
    
    # Streamlit Print
    st.dataframe(df_compare_styled)

    #================================================================================= PALLETS ===================================================================================================================================== #

    st.subheader(":two: Pallet:", divider="blue")

    # READ DSBP PALLETS #
    df_dsbp_pallet = pd.read_excel(filename_dsbp, sheet_name="Strategies & Counts")
    df_dsbp_pallet_keep = ["PI FPC Code (IL/PM, IOL)", "PI FPC Description (IL/PM, IOL)", "Intended Markets (IL/PM, IOL)", "Production Plant (Primary) Pallet (IL/PM, IOL)"]
    df_dsbp_pallet = df_dsbp_pallet.drop(df_dsbp_pallet.columns.difference(df_dsbp_pallet_keep), axis=1)
    df_dsbp_pallet = df_dsbp_pallet.rename(columns={"PI FPC Code (IL/PM, IOL)": "FPP Name", "PI FPC Description (IL/PM, IOL)": "FPP Title", "Intended Markets (IL/PM, IOL)": "Markets", "Production Plant (Primary) Pallet (IL/PM, IOL)": "Pallet Type_DSBP"})
    df_dsbp_pallet = df_dsbp_pallet.astype(str)

    # READ DSM pallet #
    df_dsm_pallet = pd.read_excel(filename_dsm, sheet_name="Weights & Dimensions")
    df_dsm_pallet = df_dsm_pallet[df_dsm_pallet['Transport Unit - Pallet Type'].notnull()]
    df_dsm_pallet = df_dsm_pallet[df_dsm_pallet['TransportUnit - Include In SAP BOM Feed'] == "Yes"]
    df_dsm_pallet_keep = ["Name/Number", "Title", "Transport Unit - Pallet Type", "Transport Unit - Name", "TransportUnit - Include In SAP BOM Feed", "Transport Unit - Stacking Pattern GCAS Code"]
    df_dsm_pallet = df_dsm_pallet.drop(df_dsm_pallet.columns.difference(df_dsm_pallet_keep), axis=1)
    df_dsm_pallet = df_dsm_pallet.rename(columns={"Name/Number": "FPP Name", "Title":"FPP Title", "Transport Unit - Pallet Type": "Pallet Type_DSM", "Transport Unit - Name": "TUP DSM", "Transport Unit - Stacking Pattern GCAS Code": "SPS DSM"})
    df_dsm_pallet = df_dsm_pallet.astype(str)

    # MERGE Pallet #
    df_merged_pallet = pd.merge(df_dsbp_pallet, df_dsm_pallet, on=['FPP Name'], how='outer')
    column_order = ['FPP Name', 'Markets', 'Pallet Type_DSBP', 'Pallet Type_DSM', 'TUP DSM', 'SPS DSM']
    df_merged_pallet = df_merged_pallet[column_order]
    df_merged_pallet = df_merged_pallet.sort_values(by='FPP Name', ascending=True)
    df_merged_pallet = df_merged_pallet.reset_index()
    df_merged_pallet.drop("index", axis=1, inplace=True)

    # Apply the highlight_diff function to each row of the DataFrame
    df_compare_styled = df_merged_pallet.style.apply(highlight_diff, axis=1)
    
    # Streamlit Print
    st.dataframe(df_compare_styled)
    
    #  ===================================================================================== PLANTS ==============================================================================
    
    st.subheader(":three: Plant:", divider="blue")
    
    # READ DSBP PLANTS #
    df_dsbp_plant = pd.read_excel(filename_dsbp, sheet_name="Production Data")
    df_dsbp_plant_keep = ["PI FPC Code (IL/PM)", "PI FPC Description (IL/PM)", "Producing Plant (Artwork planner, Business planner, C_ACE, GPS, IL/PM, IMDO, IOL, MIL, MSM, PIL, Package SPOC, Purchasing, R&D Formula, R&D Pack, SIEL, TRL, Tech Pack & MPD)"]
    df_dsbp_plant = df_dsbp_plant.drop(df_dsbp_plant.columns.difference(df_dsbp_plant_keep), axis=1)
    df_dsbp_plant = df_dsbp_plant.rename(columns={"PI FPC Code (IL/PM)": "FPP Name", "PI FPC Description (IL/PM)": "FPP Title", "Producing Plant (Artwork planner, Business planner, C_ACE, GPS, IL/PM, IMDO, IOL, MIL, MSM, PIL, Package SPOC, Purchasing, R&D Formula, R&D Pack, SIEL, TRL, Tech Pack & MPD)": "Plant_DSBP"})
    df_dsbp_plant = df_dsbp_plant.astype(str)

    # READ DSM PLANTS # 
    df_dsm_plant = pd.read_excel(filename_dsm, sheet_name="Plants")
    df_dsm_plant = df_dsm_plant[df_dsm_plant["Type"] == "Finished Product Part"]
    df_dsm_plant_keep = ["Name/Number", "Title", "Plants"]
    df_dsm_plant = df_dsm_plant.drop(df_dsm_plant.columns.difference(df_dsm_plant_keep), axis=1)
    df_dsm_plant = df_dsm_plant.rename(columns={"Name/Number": "FPP Name", "Title": "FPP Title", "Plants": "Plant_DSM"})
    df_dsm_plant = df_dsm_plant.astype(str)

    # MERGE PLANTS #
    df_merged_plant = pd.merge(df_dsbp_plant, df_dsm_plant, on=['FPP Name'], how='outer')
    column_order = ['FPP Name', 'Plant_DSBP', 'Plant_DSM']
    df_merged_plant = df_merged_plant[column_order]
    df_merged_plant = df_merged_plant.sort_values(by='FPP Name', ascending=True)
    df_merged_plant = df_merged_plant.reset_index()
    df_merged_plant.drop("index", axis=1, inplace=True)
    df_merged_plant['Plant_DSM'] = df_merged_plant['Plant_DSM'].str[:-5] #eliminates last 5 letters of the strings in that column to get rid of plant code: blablabla~D5O9 --> blablabla
    
    # Apply the highlight_diff function to each row of the DataFrame
    # df_compare_styled = df_merged_plant.style.apply(highlight_diff, axis=1)

    # Streamlit Print
    st.dataframe(df_merged_plant)


    # ================================================================================== MASTERS ======================================================================== #
    
    st.subheader(":four: TUP & SPS:", divider="blue")


    df_dsm_pallet2 = pd.read_excel(filename_dsm, sheet_name="Weights & Dimensions")
    df_dsm_pallet2 = df_dsm_pallet2[df_dsm_pallet2['Transport Unit - Pallet Type'].notnull()]
    df_dsm_pallet2 = df_dsm_pallet2[df_dsm_pallet2['TransportUnit - Include In SAP BOM Feed'] == "Yes"]
    df_dsm_pallet2_keep = ["Name/Number", "Title", "Transport Unit - Pallet Type", "Transport Unit - Name", "Transport Unit - Stacking Pattern GCAS Code", 
                           "Transport Unit - Number of Customer Units per Layer", "Transport Unit - Number of Layers per Transport Unit"]
    df_dsm_pallet2 = df_dsm_pallet2.drop(df_dsm_pallet2.columns.difference(df_dsm_pallet2_keep), axis=1)
    df_dsm_pallet2 = df_dsm_pallet2.rename(columns={"Name/Number": "FPP Name", "Title":"FPP Title", "Transport Unit - Pallet Type": "Pallet Type", "Transport Unit - Name": "TUP_DSM", 
                                                    "Transport Unit - Stacking Pattern GCAS Code": "SPS_DSM", "Transport Unit - Number of Customer Units per Layer": "Count",
                                                    "Transport Unit - Number of Layers per Transport Unit": "Layers"})
    df_dsm_pallet2 = df_dsm_pallet2.astype(str)

    df_tracksheet = pd.read_excel(filename_tracksheet, sheet_name="MBOM")
    df_tracksheet_keep = ["Execution", "Pallet type", "Count", "Layers", "SPS", "TUP"]
    df_tracksheet = df_tracksheet.drop(df_tracksheet.columns.difference(df_tracksheet_keep), axis=1)
    df_tracksheet = df_tracksheet.rename(columns={"Execution": "Execution (Tracksheet)", "Pallet type": "Pallet Type", "Count":"Count", "Layers":"Layers", "SPS":"SPS_TS", "TUP":"TUP_TS"})
    df_tracksheet = df_tracksheet.astype(str)

    df_merged_tracksheet = pd.merge(df_tracksheet, df_dsm_pallet2, on=['Pallet Type', 'Count', 'Layers'], how='outer')
    column_order = ['FPP Name', 'FPP Title', 'Execution (Tracksheet)', 'Pallet Type', 'Count', 'Layers', 'TUP_TS', 'TUP_DSM', 'SPS_TS', 'SPS_DSM']
    df_merged_tracksheet = df_merged_tracksheet[column_order]
    df_merged_tracksheet = df_merged_tracksheet.dropna(subset=['FPP Name'])
    df_merged_tracksheet = df_merged_tracksheet.sort_values(by='FPP Name', ascending=True)
    df_merged_tracksheet = df_merged_tracksheet.reset_index()
    df_merged_tracksheet.drop("index", axis=1, inplace=True)



    def highlight_diff2(row):
        color_map = []
        for col in row.index:
            if col.endswith("_TS") or col.endswith("_DSM"):
                col_ts = col.replace("_DSM", "_TS")
                col_dsm = col.replace("_TS", "_DSM")
                if row.get(col_ts) != row.get(col_dsm):
                    color_map.append("background-color: lightcoral")
                else:
                    color_map.append("background-color: lightgreen")
            else:
                color_map.append("")
        return color_map

    # Apply the highlight_diff function to each row of the DataFrame
    df_compare_styled = df_merged_tracksheet.style.apply(highlight_diff2, axis=1)
    
    # Streamlit Print
    st.dataframe(df_compare_styled)
    


# Main function
def main():
    with st.sidebar:
        selected = option_menu(
            menu_title="Navigation",
            options=["SpecProof"],
            icons=["robot"],
            menu_icon="list",
            default_index=0,
        )

    # Render the selected page
    if selected == "SpecProof":
        SPECPROOF()

if __name__ == "__main__":
    main()
