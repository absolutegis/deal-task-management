import pandas as pd
import streamlit as st
import re

# Set the layout to wide
st.set_page_config(layout="wide")

# Add a logo and subtext
logo_path = "https://storage.googleapis.com/absolute_gis_public/Images/lennar_indy.jpg"  # Replace with your logo file name or URL
st.image(logo_path, width=150)
st.markdown("<h4 style='text-align: left; color: #015cab;'>Created by Lennar Indianapolis Division</h4>", unsafe_allow_html=True)

# Function to clean up the column names
def clean_column_names(columns):
    return [re.sub(r'\s*\(.*?\)', '', col) for col in columns]

# Load the Excel file and identify the sheet
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
if uploaded_file is not None:
    sheet_name = pd.ExcelFile(uploaded_file).sheet_names[0]
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    
    # Clean the column names first to ensure consistent access
    df.columns = clean_column_names(df.columns)
    
    # Define the expected columns for deals and tasks
    deal_columns = [
        'Regarding', 'Sub-Market', 'Calculated Deal Stage', 
        'GF Submittal Date', 'Green Folder Meeting Date', 
        'IP Expiration Date', 'Days to IP Expiration', 
        'Projected Deal First Closing Date', 'Deal Homesite Total', 
        'Homesite Size Description', 'Acquisition Type', 
        'Primary Seller Company', 'Product Type Description'
    ]
    
    task_columns = [
        'Task Category', 'Owner', 'Subject', 'Start Date', 'Due Date', 
        'Vendor Assigned', 'Priority', 'Comment', 'Status Reason'
    ]
    
    # Identify columns in the uploaded file
    available_columns = df.columns
    
    # Extract relevant columns for deals and tasks
    deals_df = df[[col for col in clean_column_names(deal_columns) if col in available_columns]].drop_duplicates()
    tasks_df = df[[col for col in clean_column_names(task_columns) if col in available_columns] + ['Regarding']].drop_duplicates()

    # Streamlit Interface
    st.title("Deal and Task Management Interface")

    # Dropdown for selecting a deal
    deal_selected = st.selectbox("Select a Deal", ["Show All"] + list(deals_df['Regarding']), index=0)

    # Reset Button
    if st.button('Reset View'):
        deal_selected = "Show All"

    # Display all deals or scroll to the selected deal
    for _, deal in deals_df.iterrows():
        deal_id = deal['Regarding'].replace(' ', '_')
        st.markdown(f"<div id='{deal_id}'></div>", unsafe_allow_html=True)
        
        if deal_selected == "Show All" or deal['Regarding'] == deal_selected:
            st.subheader(f"Deal: {deal['Regarding']}")
            st.table(deal[clean_column_names(deal_columns)].to_frame().T)
            
            related_tasks = tasks_df[tasks_df['Regarding'] == deal['Regarding']]
            
            with st.expander("Related Tasks", expanded=True):
                if not related_tasks.empty:
                    def color_status(val):
                        if val == 'Completed':
                            return 'background-color: gray'
                        elif val == 'In Progress':
                            return 'background-color: green'
                        elif pd.to_datetime(val, errors='coerce') < pd.to_datetime('today'):
                            return 'background-color: red'
                        else:
                            return ''
                    
                    styled_tasks = related_tasks.style.applymap(
                        color_status, subset=['Status Reason', 'Due Date']
                    ).set_properties(**{'text-align': 'left', 'width': '100%'})
                    
                    st.dataframe(styled_tasks)
                else:
                    st.write("No related tasks found.")
    
    # Scroll to the selected deal
    if deal_selected != "Show All":
        st.markdown(f"""
            <script>
                var dealElement = document.getElementById('{deal_selected.replace(' ', '_')}');
                if (dealElement) {{
                    dealElement.scrollIntoView({{behavior: 'smooth', block: 'center'}});
                }}
            </script>
        """, unsafe_allow_html=True)