import pandas as pd
import streamlit as st
import re
from bs4 import BeautifulSoup
from io import BytesIO
from datetime import datetime
import plotly.express as px

# Set the layout to wide
st.set_page_config(layout="wide")

# Create an anchor at the top of the page
st.markdown('<a name="top"></a>', unsafe_allow_html=True)

# Add a logo and subtext
logo_path = "https://storage.googleapis.com/absolute_gis_public/Images/lennar_indy.jpg"  # Replace with your logo file name or URL
st.image(logo_path, width=300)

# Move the bold text 'Deal and Task Management Interface' here
st.title("Deal and Task Management Interface ")

# Add the subtext after the title
st.markdown(
    "<h4 style='text-align: left; color: #015cab;'>Created by Lennar Indianapolis Division</h4>",
    unsafe_allow_html=True
)

# Inject custom CSS for buttons via HTML
st.markdown(
    """
    <style>
    div.stButton > button:first-child {
        background-color: #015cab;
        color: white;
        padding: 0.5em 1em;
        border: none;
        border-radius: 5px;
        cursor: pointer;
        transition: background-color 0.3s, color 0.3s;
    }
    div.stButton > button:first-child:hover {
        background-color: #e04e2b;
    }
    div.stButton > button:first-child:active {
        background-color: #e04e2b !important;
        color: white !important;
    }
    div.stButton > button:focus {
        outline: none;
        box-shadow: none;
        background-color: #FF5733 !important;
        color: white !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Function to clean up the column names by stripping out '(Regarding) (Deal)'
def clean_column_names(columns):
    return [re.sub(r'\s*\(.*?\)', '', col).strip() for col in columns] 

# Function to strip HTML tags and retain only text
def strip_html(text):
    if isinstance(text, str):
        soup = BeautifulSoup(text, "html.parser")
        return soup.get_text(separator=" ", strip=True)
    return text

# Function to apply conditional formatting with semi-transparency
def apply_conditional_formatting(styled_df):
    styled_df = styled_df.applymap(lambda val: 'background-color: rgba(255, 0, 0, 0.3)' if val == 'Overdue' else '', subset=['Status Reason', 'Due Date'])
    styled_df = styled_df.applymap(lambda val: 'background-color: rgba(0, 128, 0, 0.3)' if val == 'In Progress' else '', subset=['Status Reason'])
    styled_df = styled_df.applymap(lambda val: 'background-color: rgba(128, 128, 128, 0.3)' if val == 'Completed' else '', subset=['Status Reason'])
    return styled_df

# Function to generate the Gantt chart
def generate_gantt_chart(deal_name, deal, filtered_tasks_df):
    gantt_data = {
        'Task': ['GF Submittal Date', 'Green Folder Meeting Date', 'IP Expiration Date', 'Projected Deal First Closing Date'],
        'Start': [deal['GF Submittal Date'], deal['Green Folder Meeting Date'], deal['IP Expiration Date'], deal['Projected Deal First Closing Date']],
        'Finish': [deal['Green Folder Meeting Date'], deal['IP Expiration Date'], deal['Projected Deal First Closing Date'], deal['Projected Deal First Closing Date']]
    }

    # Adding Tasks Start Date and Tasks Due Date
    tasks_subset = filtered_tasks_df[['Subject', 'Start Date', 'Due Date']].dropna()
    for _, task in tasks_subset.iterrows():
        gantt_data['Task'].append(task['Subject'])
        gantt_data['Start'].append(task['Start Date'])
        gantt_data['Finish'].append(task['Due Date'])

    gantt_df = pd.DataFrame(gantt_data)

    fig = px.timeline(
        gantt_df,
        x_start="Start",
        x_end="Finish",
        y="Task",
        title=f"Gantt Chart for {deal_name}",
        color="Task"
    )

    fig.update_yaxes(categoryorder="total ascending")
    fig.update_layout(showlegend=False)
    return fig

# Load the Excel files
uploaded_files = st.file_uploader(
    "Choose Excel files",
    type="xlsx",
    accept_multiple_files=True,
    help="Upload two Excel files: one for Deals/Tasks and one for Appointments"
)

if uploaded_files and len(uploaded_files) == 2:
    try:
        # Initialize Excel writer
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            current_row = 0  # Initialize starting row for Excel

            # Read the uploaded files into dataframes
            df1 = pd.read_excel(uploaded_files[0])
            df2 = pd.read_excel(uploaded_files[1])

            # Clean the column names
            df1.columns = clean_column_names(df1.columns)
            df2.columns = clean_column_names(df2.columns)

            # Identify which dataframe is appointments based on specific columns
            if {'Subject', 'Start Time'}.issubset(df1.columns):
                appointments_df = df1.copy()
                deals_tasks_df = df2.copy()
            elif {'Subject', 'Start Time'}.issubset(df2.columns):
                appointments_df = df2.copy()
                deals_tasks_df = df1.copy()
            else:
                st.error("Could not identify the Appointments file. Please ensure it contains 'Subject' and 'Start Time' columns.")
                st.stop()

            # Clean 'Description' field in appointments
            if 'Description' in appointments_df.columns:
                appointments_df['Description'] = appointments_df['Description'].apply(strip_html)

            # Drop unwanted columns from appointments
            appointments_df = appointments_df.drop(columns=['Appointment', 'Row Checksum', 'Modified On'], errors='ignore')

            # Define expected columns (after cleaning)
            deal_columns = [
                'Regarding', 'Sub-Market', 'Calculated Deal Stage',
                'GF Submittal Date', 'Green Folder Meeting Date',
                'IP Expiration Date', 'Days to IP Expiration',
                'Projected Deal First Closing Date', 'Deal Homesite Total',
                'Homesite Size Description', 'Acquisition Type',
                'Primary Seller Company', 'Product Type Description',
                'CIC Final Approval Date'
            ]

            task_columns = [
                'Task Category', 'Owner', 'Subject', 'Start Date', 'Due Date',
                'Vendor Assigned', 'Priority', 'Comment', 'Status Reason'
            ]

            appointment_columns = [
                'Subject', 'Regarding', 'Owner', 'Status', 'Start Time',
                'End Time', 'Category', 'Description'
            ]

            # Extract deals data
            deals_df = deals_tasks_df[deal_columns].drop_duplicates().reset_index(drop=True)

            # Handle non-finite values in 'Days to IP Expiration'
            deals_df['Days to IP Expiration'] = deals_df['Days to IP Expiration'].fillna(0).round().astype(int)
            
            # Convert date fields to datetime and extract only the date
            date_columns = ['GF Submittal Date', 'Green Folder Meeting Date', 'IP Expiration Date', 'Projected Deal First Closing Date', 'CIC Final Approval Date']
            deals_df[date_columns] = deals_df[date_columns].apply(lambda x: pd.to_datetime(x, errors='coerce').dt.date)

            # Extract tasks data
            tasks_df = deals_tasks_df[task_columns + ['Regarding']].drop_duplicates().reset_index(drop=True)
            tasks_df[['Start Date', 'Due Date']] = tasks_df[['Start Date', 'Due Date']].apply(lambda x: pd.to_datetime(x, errors='coerce').dt.date)

            # Extract and clean appointments data
            appointments_df[['Start Time', 'End Time']] = appointments_df[['Start Time', 'End Time']].apply(lambda x: pd.to_datetime(x, errors='coerce').dt.date)

            # Deal Filters

            # Greenfolder Approved, Not Yet Closed
            greenfolder_approved_df = deals_df[deals_df['CIC Final Approval Date'].notna()]

            # Green Folder Schedule
            green_folder_schedule_df = deals_df[
                (deals_df['GF Submittal Date'].notna()) &
                (deals_df['CIC Final Approval Date'].isna())
            ]

            # Letters of Intent
            letters_of_intent_df = deals_df[
                deals_df['GF Submittal Date'].isna()
            ].sort_values(by=["Calculated Deal Stage", "Sub-Market"])

            # Add buttons for Deal Filters with counts
            # Add buttons for Deal Filters with counts
            col1, col2, col3, col4 = st.columns(4)

            with col1:
                if st.button(f"Greenfolder Approved, Not Yet Closed ({len(greenfolder_approved_df)})"):
                    st.session_state['deal_filter'] = 'Greenfolder Approved, Not Yet Closed'
                    st.session_state['filtered_deals_df'] = greenfolder_approved_df
            with col2:
                if st.button(f"Green Folder Schedule ({len(green_folder_schedule_df)})"):
                    st.session_state['deal_filter'] = 'Green Folder Schedule'
                    st.session_state['filtered_deals_df'] = green_folder_schedule_df
            with col3:
                if st.button(f"Letters of Intent ({len(letters_of_intent_df)})"):
                    st.session_state['deal_filter'] = 'Letters of Intent'
                    st.session_state['filtered_deals_df'] = letters_of_intent_df
            with col4:
                total_deals_count = len(deals_df)
                if st.button(f"Reset Filter ({total_deals_count})"):
                    st.session_state.pop('filtered_deals_df', None)  # Remove the filtered deals to reset to all deals
                    st.session_state.pop('deal_filter', None)  # Clear the deal filter state as well

            # Default to showing all deals if no button is clicked
            if 'filtered_deals_df' not in st.session_state:
                st.session_state['filtered_deals_df'] = deals_df

            filtered_deals_df = st.session_state['filtered_deals_df']


            # Sorting UI/UX
            with st.expander("Sort Deals"):
                sort_column = st.selectbox(
                    "Sort by:",
                    options=["Projected Deal First Closing Date", "GF Submittal Date", "Calculated Deal Stage", "Sub-Market"],
                    index=0
                )

                sort_order = st.radio(
                    "Sort Order:",
                    options=["Ascending", "Descending"],
                    index=0,
                    horizontal=True
                )

                if sort_order == "Ascending":
                    filtered_deals_df = filtered_deals_df.sort_values(by=sort_column, ascending=True)
                else:
                    filtered_deals_df = filtered_deals_df.sort_values(by=sort_column, ascending=False)

            # Search Functionality using Dropdown with Search
            with st.expander("Search for Specific Deal"):
                deal_names = filtered_deals_df['Regarding'].dropna().unique().tolist()
                selected_deal = st.selectbox("Select a Deal:", [""] + deal_names)
                if selected_deal:
                    filtered_deals_df = filtered_deals_df[filtered_deals_df['Regarding'] == selected_deal]

            # Add buttons to minimize/maximize all tasks and appointments
            col1, col2, col3 = st.columns([1, 1, 1])

            with col1:
                minimize_all_button = st.button("Minimize All")
            with col2:
                maximize_all_button = st.button("Maximize All")
                
            # Initialize expander states at the start
            expander_states = {}

            # Minimize/Maximize all tasks and appointments
            if minimize_all_button:
                expander_states = {deal: False for deal in filtered_deals_df['Regarding']}
            if maximize_all_button:
                expander_states = {deal: True for deal in filtered_deals_df['Regarding']}

            # Loop through filtered deals
            for idx, deal in filtered_deals_df.iterrows():
                deal_name = deal['Regarding']

                # Add a 'Return to Top' link next to the deal name with a home emoji
                return_to_top_link = f"<a href='#top' style='text-decoration: none; color: #015CAB;'>🏠</a>"
                st.markdown(f"<h3>Deal: {deal_name} {return_to_top_link}</h3>", unsafe_allow_html=True)

                # Display Deal Data
                deal_data = deal.to_frame().T
                st.table(deal_data)

                # Write Deal Data to Excel
                deal_data.to_excel(writer, startrow=current_row, index=False, header=True, sheet_name='Data')
                current_row += len(deal_data) + 2  # Adjust row position

                # Fetch and display related Tasks
                filtered_tasks_df = tasks_df[tasks_df['Regarding'] == deal_name]

                # Count the number of tasks per status
                in_progress_count = filtered_tasks_df[filtered_tasks_df['Status Reason'] == 'In Progress'].shape[0]
                completed_count = filtered_tasks_df[filtered_tasks_df['Status Reason'] == 'Completed'].shape[0]
                not_started_count = filtered_tasks_df[filtered_tasks_df['Status Reason'] == 'Not Started'].shape[0]
                total_tasks_count = len(filtered_tasks_df)

                # Initialize task_filter with a default value
                task_filter = "Show All"  # Default filter value

                # Create a row of buttons for filtering tasks with unique keys
                col1, col2, col3, col4 = st.columns(4)

                with col1:
                    if st.button(f"Show All Tasks ({total_tasks_count})", key=f"{deal_name}_show_all_{idx}"):
                        task_filter = "Show All"
                with col2:
                    if st.button(f"In Progress ({in_progress_count})", key=f"{deal_name}_in_progress_{idx}"):
                        task_filter = "In Progress"
                with col3:
                    if st.button(f"Completed ({completed_count})", key=f"{deal_name}_completed_{idx}"):
                        task_filter = "Completed"
                with col4:
                    if st.button(f"Not Started ({not_started_count})", key=f"{deal_name}_not_started_{idx}"):
                        task_filter = "Not Started"

                # Filter the DataFrame based on the button clicked
                if task_filter != "Show All":
                    filtered_tasks_df = filtered_tasks_df[filtered_tasks_df['Status Reason'] == task_filter]

                # Construct the label for the expander
                task_count = len(filtered_tasks_df)
                expander_label = f"Related Tasks ({task_count})"

                # Display the tasks in an expander
                with st.expander(expander_label, expanded=expander_states.get(deal_name, False)):
                    if not filtered_tasks_df.empty:
                        styled_tasks = apply_conditional_formatting(filtered_tasks_df.drop(columns=['Regarding']).style)
                        st.dataframe(styled_tasks)  # Let Streamlit automatically determine the height
                        # Write Tasks Data to Excel
                        filtered_tasks_df.drop(columns=['Regarding']).to_excel(writer, startrow=current_row, index=False, header=True, sheet_name='Data')
                        current_row += len(filtered_tasks_df) + 2
                    else:
                        st.write("No related tasks found.")
                        current_row += 2  # Add spacing even if no tasks

                # Fetch and display related Appointments
                related_appointments = appointments_df[appointments_df['Regarding'] == deal_name].drop(columns=['Regarding'])

                # Calculate the number of related appointments
                appointment_count = len(related_appointments)

                # Display the appointments in an expander
                with st.expander(f"Related Appointments ({appointment_count})", expanded=expander_states.get(deal_name, False)):
                    if not related_appointments.empty:
                        st.dataframe(related_appointments)  # Let Streamlit automatically determine the height
                        # Write Appointments Data to Excel
                        related_appointments.to_excel(writer, startrow=current_row, index=False, header=True, sheet_name='Data')
                        current_row += len(related_appointments) + 2
                    else:
                        st.write("No related appointments found.")
                        current_row += 2  # Add spacing even if no appointments

                # Gantt chart generation button using Streamlit with custom styling
                if st.button(f"Generate Gantt Chart for {deal_name}", key=f"gantt_{idx}_{deal_name}"):
                    fig = generate_gantt_chart(deal_name, deal, filtered_tasks_df)
                    st.plotly_chart(fig)

                # Add a more prominent separator row
                st.markdown("<hr style='border: 4px solid #000;'>", unsafe_allow_html=True)  # Thicker horizontal line
                current_row += 1  # Extra space between deals

            # Adjust column widths for better readability
            worksheet = writer.sheets['Data']
            for i, col in enumerate(deal_columns):  # Adjust for deal columns
                worksheet.set_column(i, i, 20)
            for i, col in enumerate(task_columns):  # Adjust for task columns
                worksheet.set_column(i, i, 20)
            for i, col in enumerate(appointment_columns):  # Adjust for appointment columns
                worksheet.set_column(i, i, 20)

        # Render the download button after all the writing is done
        st.download_button(
            label="Download Excel",
            data=buffer.getvalue(),
            file_name=f"deal_task_appointment_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("Please upload exactly two Excel files: one for Deals/Tasks and one for Appointments.")
