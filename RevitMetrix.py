#libraries
import streamlit as st
import pandas as pd
import os
import warnings
import openpyxl as plx
from pandas import ExcelWriter
from openpyxl.drawing.image import Image
import io
import matplotlib
matplotlib.use('agg')
import matplotlib.pyplot as plt
warnings.filterwarnings("ignore")

#header
# pd.set_option('display.show_dimensions', False)  # Hide the dimensions display (number of rows and columns)
st.image("volantis.png")
st.title("RevitMetrix: Insightful Data Analysis voor Uittrekstaten")
uploaded_file = st.file_uploader("Kies een bestand", type=["xlsx"])

# Process the uploaded file
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("Origineel Bestand:")

    # Display nlsfb dmv assembly code en assembly description  
    df.insert(1, 'NLSfB', df['Assembly Code'] + ' ' + df['Assembly Description'])

    #laat assembly code en assembly description vallen
    columns_to_drop = ['Assembly Code', 'Assembly Description']
    df.drop(columns=columns_to_drop, inplace=True)
    
    # Convert 'Material: Unit Weight' to kg/m³
    df['Material: Unit weight [kg/m³]'] = df['Material: Unit weight'].str.replace(' kN/m³', '').str.replace(',', '.').astype(float) * 9.81

    # Convert stripped values in 'Material: Area' and 'Material: Volume' columns to float Remove commas and units, and convert to float
    df['Material: Area'] = df['Material: Area'].str.replace('m²', '').str.replace(',', '.').astype(float)
    df['Material: Volume'] = df['Material: Volume'].str.replace('m³', '').str.replace(',', '.').astype(float)
    df['Material: Unit weight'].str.replace('kN/m³', '').str.replace(',', '.').astype(float)
    st.dataframe(df, hide_index=True)

    # Calculate the new column by multiplying 'Material: Unit Weight' and 'Material: Volume' see l33
    df['Material: Weight'] = df['Material: Unit weight [kg/m³]'] * df['Material: Volume']

    # Group by 'Material: Name' and calculate the sum of 'Material: Volume' and 'Material: Area'
    grouped_df_familyandtype = df.groupby(['Family and Type', 'Material: Unit weight [kg/m³]'])[['Family and Type', 'Count', 'Material: Volume', 'Material: Area', 'Material: Weight']].sum()
    grouped_df_familyandtype = grouped_df_familyandtype.sort_values(by='Material: Weight', ascending=False)
    grouped_df_familyandtype = grouped_df_familyandtype.reset_index(drop=True)
    st.write("Family and Type")
    # print(grouped_df_familyandtype.columns)
    st.dataframe(grouped_df_familyandtype, hide_index=True)

    # Group by 'Material: Name' and calculate the sum of 'Material: Volume' and 'Material: Area'
    grouped_df_materialname = df.groupby(['Material: Name', 'Material: Unit weight [kg/m³]'])[['Material: Name', 'Count','Material: Volume', 'Material: Area', 'Material: Weight']].sum()
    grouped_df_materialname = grouped_df_materialname.sort_values(by='Material: Weight', ascending=False)
    grouped_df_materialname = grouped_df_materialname.reset_index(drop=True)
    st.write("Material: Name")
    # print(grouped_df_materialname.columns)
    st.dataframe(grouped_df_materialname, hide_index=True)
    
    # Group by 'Description' and calculate the sum of 'Material: Volume' and 'Material: Area'
    grouped_df_description = df.groupby(['Description', 'Material: Unit weight [kg/m³]'])[['Description', 'Count','Material: Volume', 'Material: Area', 'Material: Weight']].sum()
    grouped_df_description = grouped_df_description.sort_values(by='Material: Weight', ascending=False)
    grouped_df_description = grouped_df_description.reset_index(drop=True)
    st.write('Descripton:')
    # print(grouped_df_description.columns)
    st.dataframe(grouped_df_description, hide_index=True)

    # Group by 'nlsfb' and calculate the sum of 'Material: Volume' and 'Material: Area'
    grouped_df_nlsfb = df.groupby(['NLSfB', 'Material: Unit weight [kg/m³]'])[['NLSfB', 'Count','Material: Volume', 'Material: Area', 'Material: Weight']].sum()
    grouped_df_nlsfb = grouped_df_nlsfb.sort_values(by='Material: Weight', ascending=False)
    grouped_df_nlsfb = grouped_df_nlsfb.reset_index(drop=True)
    st.write('NLSfB:')
    # print(grouped_df_nlsfb.columns)
    st.dataframe(grouped_df_nlsfb, hide_index=True)

    # Let the user input the desired Excel file name
    user_entered_name = st.text_input("Kies een naam voor het Excel-bestand:", "Projectnaam.xlsx")
    excel_file_name = user_entered_name.strip()  # Remove leading/trailing spaces

    # Display a link to download the Excel file
    if st.button("Download Excel-bestand"):
        
        # Define the directory where you want to save the Excel file
        save_directory = os.path.join(os.path.expanduser("~"), "Downloads")

        # Create the directory if it doesn't exist
        os.makedirs(save_directory, exist_ok=True)

        # Define the full path for the Excel file
        excel_file_path = os.path.join(save_directory, excel_file_name)

        # Export the summarized DataFrames to separate sheets in the same Excel file
        with ExcelWriter(excel_file_path) as writer:
            df.to_excel(writer, sheet_name="Origineel", index=False)
            grouped_df_materialname.to_excel(writer, sheet_name='Material Name', index=False)
            grouped_df_familyandtype.to_excel(writer, sheet_name='Family and Type', index=False)
            grouped_df_description.to_excel(writer, sheet_name="Description", index=False)
            grouped_df_nlsfb.to_excel(writer, sheet_name="NLSfB", index=False)

            # Create bar charts and embed them in the Excel file
            sheet_materialname = writer.sheets['Material Name']
            sheet_familyandtype = writer.sheets['Family and Type']
            sheet_description = writer.sheets['Description']
            sheet_nlsfb = writer.sheets['NLSfB']

            # Create bar chart for Family and Type
            plt.figure(figsize=(20, 12))
            grouped_df_familyandtype.plot(kind='bar', x='Family and Type', y='Material: Weight')
            plt.title('Family and Type | Material: Weight [kg]')
            plt.xlabel('Family and Type')
            plt.ylabel('Material: Weight [kg]')
            img_data = io.BytesIO()
            plt.savefig(img_data, format='png')
            img = Image(img_data)
            sheet_familyandtype.add_image(img, 'H5')

            # Create bar chart for Material: Name
            plt.figure(figsize=(20, 12))
            grouped_df_materialname.plot(kind='bar', x='Material: Name', y='Material: Weight')
            plt.title('Material: Weight | Material: Weight [kg]')
            plt.xlabel('Material: Name')
            plt.ylabel('Material: Weight [kg]')
            plt.xticks(rotation=45)
            img_data = io.BytesIO()
            plt.savefig(img_data, format='png')
            img = Image(img_data)
            sheet_materialname.add_image(img, 'H5')

            # Create bar chart for Description
            plt.figure(figsize=(20, 12))
            grouped_df_description.plot(kind='bar', x='Description', y='Material: Weight')
            plt.title('Description Bar Chart | Material: Weight [kg]')
            plt.xlabel('Description')
            plt.ylabel('Material: Weight [kg]')
            img_data = io.BytesIO()
            plt.savefig(img_data, format='png')
            img = Image(img_data)
            sheet_description.add_image(img, 'H5')

            # Create bar chart for nlsfb
            plt.figure(figsize=(20, 12))
            grouped_df_nlsfb.plot(kind='bar', x='NLSfB', y='Material: Weight')
            plt.title('NL-SfB Bar Chart | Material: Weight [kg]')
            plt.xlabel('NL-SfB')
            plt.ylabel('Material: Weight [kg]')
            img_data = io.BytesIO()
            plt.savefig(img_data, format='png')
            img = Image(img_data)
            sheet_nlsfb.add_image(img, 'H5')

        # Check if the Excel file exists at the given path
        if os.path.exists(excel_file_path):
            st.write("Klik hier om het Excel-bestand te downloaden:")
            st.download_button(
                label=excel_file_name,
                data=open(excel_file_path, "rb").read(),
                file_name=excel_file_name,
                key="download-button"
            )
        else:
            st.warning("Het Excel-bestand is nog niet gegenereerd. Upload een bestand en klik opnieuw op de knop.")
