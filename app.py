import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# Streamlit app title
st.title("Excel Processor")

# Step 1: Upload Excel file
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

# Step 2: Date selection for subtraction
selected_date = st.date_input("Select a date to subtract from Date of Disbursement")

if uploaded_file:
    # Step 3: Read all sheets from the uploaded Excel file
    excel_df = pd.read_excel(uploaded_file, sheet_name=None)

    # Initialize an empty DataFrame for processing data
    combined_df = pd.DataFrame()

    # Step 4: Flatten sheets and process the data
    for sheet_name, sheet_df in excel_df.items():
        # Ensure 'Date of Disbursement' is in datetime format
        if 'Date of Disbursement' in sheet_df.columns:
            sheet_df['Date of Disbursement'] = pd.to_datetime(sheet_df['Date of Disbursement'], errors='coerce')

        # Process the sheet and add it to the combined DataFrame
        combined_df = pd.concat([combined_df, sheet_df], ignore_index=True)

    # Step 5: Process the DataFrame if 'Date of Disbursement' exists
    if 'Date of Disbursement' in combined_df.columns:
        # Convert selected_date to datetime
        selected_date = pd.to_datetime(selected_date)

        # Calculate 'new_ageing'
        combined_df['new_ageing'] = (selected_date - combined_df['Date of Disbursement']).dt.days

        # Handle missing or invalid 'Date of Disbursement'
        if combined_df['new_ageing'].isna().sum() > 0:
            st.warning(f"Warning: There are {combined_df['new_ageing'].isna().sum()} rows with invalid 'Date of Disbursement'.")
            st.write(combined_df[combined_df['new_ageing'].isna()])

        # Define slab conditions
        conditions = [
            (combined_df['new_ageing'] <= 60),
            (combined_df['new_ageing'] > 60) & (combined_df['new_ageing'] <= 90),
            (combined_df['new_ageing'] > 90) & (combined_df['new_ageing'] <= 180),
            (combined_df['new_ageing'] > 180) & (combined_df['new_ageing'] <= 365),
            (combined_df['new_ageing'] > 365)
        ]
        slab_values = ['<=60', '>60', '>90', '>180', '>365']

        # Assign "new_slab" based on conditions
        combined_df['new_slab'] = 'No Slab'  # Default value
        for i, condition in enumerate(conditions):
            combined_df.loc[condition, 'new_slab'] = slab_values[i]

        # Update existing columns if they exist
        if 'Ageing' in combined_df.columns:
            combined_df['Ageing'] = combined_df['new_ageing']
        if 'Slab' in combined_df.columns:
            combined_df['Slab'] = combined_df['new_slab']

        # Debugging: Display updated columns
        st.write("Columns after updating Ageing and Slab:")
        st.write(combined_df[['Ageing', 'Slab', 'new_ageing', 'new_slab']].head())

    else:
        st.error("'Date of Disbursement' column not found in the file.")

    # Step 6: Save the processed DataFrame to an Excel file
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        combined_df.to_excel(writer, index=False, sheet_name='Processed Data')

    # Step 7: Provide the download link
    st.success("The Excel file has been successfully processed!")
    st.download_button(
        label="Download Processed Excel File",
        data=output.getvalue(),
        file_name="processed_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
