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

st.image("volantis.png")
st.title("RevitMetrix: Insightful Data Analysis voor Uittrekstaten")
uploaded_file = st.file_uploader("Kies een bestand", type=["xlsx"])

# Process the uploaded file
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("Origineel Bestand:")
    st.write(df)

    # Group by 'Material: Name' and calculate the sum of 'Material: Volume' and 'Material: Area'
    grouped_df_familyandtype = df.groupby('Family and Type')[['Count', 'Material: Volume', 'Material: Area']].sum()
    grouped_df_familyandtype = grouped_df_familyandtype.sort_values(by='Material: Volume', ascending=False)
    st.write("Family and Type")
    st.write(grouped_df_familyandtype)

    grouped_df_materialname = df.groupby('Material: Name')[['Count','Material: Volume', 'Material: Area']].sum()
    grouped_df_materialname = grouped_df_materialname.sort_values(by='Material: Volume', ascending=False)
    st.write("Material: Name")
    st.write(grouped_df_materialname)

    # Let the user input the desired Excel file name
    user_entered_name = st.text_input("Kies een naam voor het Excel-bestand:", "[Projectnaam123456789].xlsx")
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
            grouped_df_materialname.to_excel(writer, sheet_name='Material Name', index=True)
            grouped_df_familyandtype.to_excel(writer, sheet_name='Family and Type', index=True)

            # Create bar charts and embed them in the Excel file
            # sheet_origineel = writer.sheets['Origineel']
            sheet_materialname = writer.sheets['Material Name']
            sheet_familyandtype = writer.sheets['Family and Type']

            # Create bar chart for Material Name
            plt.figure(figsize=(10, 6))
            grouped_df_materialname.plot(kind='bar')
            plt.title('Material: Name Bar Chart')
            plt.xlabel('Material: Name')
            plt.ylabel('Sum')
            img_data = io.BytesIO()
            plt.savefig(img_data, format='png')
            img = Image(img_data)
            sheet_materialname.add_image(img, 'E5')

            # Create bar chart for Family and Type
            plt.figure(figsize=(10, 6))
            grouped_df_familyandtype.plot(kind='bar')
            plt.title('Family and Type Bar Chart')
            plt.xlabel('Family and Type')
            plt.ylabel('Sum')
            img_data = io.BytesIO()
            plt.savefig(img_data, format='png')
            img = Image(img_data)
            sheet_familyandtype.add_image(img, 'E5')

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
