import pandas as pd
import streamlit as st
import re
from bs4 import BeautifulSoup
from io import BytesIO
from datetime import datetime, timedelta
import plotly.express as px

# Set the layout to wide
st.set_page_config(layout="wide")

# Create an anchor at the top of the page
st.markdown('<a name="top"></a>', unsafe_allow_html=True)

# Add a logo and subtext
logo_path = "https://storage.googleapis.com/absolute_gis_public/Images/lennar%20dashboard%20title.jpg"  # Replace with your logo file name or URL
st.image(logo_path, width=600)

# Move the bold text 'Deal and Task Management Interface' here
#st.title("Deal and Task Management Interface ")

# Add the subtext after the title
#st.markdown(
#    "<h4 style='text-align: left; color: #015cab;'>Created by Lennar Indianapolis Division</h4>",
#    unsafe_allow_html=True
#)

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
        margin-right: 0.5em;  /* Adjust this value to decrease horizontal space */
        margin-left: 0.5em;   /* Adjust this value to decrease horizontal space */
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
    # Ensure unique columns before applying styling
    if styled_df.columns.duplicated().any():
        styled_df = styled_df.loc[:, ~styled_df.columns.duplicated()]

    styled_df = styled_df.map(lambda val: 'background-color: rgba(255, 0, 0, 0.3)' if val == 'Overdue' else '', subset=['Status Reason', 'Due Date'])
    styled_df = styled_df.map(lambda val: 'background-color: rgba(0, 128, 0, 0.3)' if val == 'In Progress' else '', subset=['Status Reason'])
    styled_df = styled_df.map(lambda val: 'background-color: rgba(128, 128, 128, 0.3)' if val == 'Completed' else '', subset=['Status Reason'])
    return styled_df

# Gantt chart generation function
# Gantt chart generation function
def generate_gantt_chart(deal_name, deal, filtered_tasks_df):
    current_date = pd.Timestamp(datetime.now().date())

    gantt_data = {
        'Task': [],
        'Start': [],
        'Finish': [],
        'Status': [],
        'Color': [],
    }

    # Contract Dates (Teal bar with IP Expiration as a hash mark)
    contract_start = pd.to_datetime(deal.get('Actual Contract Execution Date'), errors='coerce')
    contract_end = pd.to_datetime(deal.get('Projected Deal First Closing Date'), errors='coerce')
    ip_expiration = pd.to_datetime(deal.get('IP Expiration Date'), errors='coerce')

    # Ensure that contract dates are displayed at the top
    if pd.notna(contract_start) and pd.notna(contract_end):
        gantt_data['Task'].append('Contract Dates')
        gantt_data['Start'].append(contract_start)
        gantt_data['Finish'].append(contract_end)
        gantt_data['Status'].append('Contract')
        gantt_data['Color'].append('teal')

    # Add "Actual Contract Execution Date" as a separate row
    if pd.notna(contract_start):
        gantt_data['Task'].append('Actual Contract Execution Date')
        gantt_data['Start'].append(contract_start)
        gantt_data['Finish'].append(contract_start)  # Start and finish on the same day
        gantt_data['Status'].append('Actual Contract Execution')
        gantt_data['Color'].append('blue')

    # Green Folder Dates (Purple bar)
    green_start = pd.to_datetime(deal.get('GF Submittal Date'), errors='coerce')
    green_end = pd.to_datetime(deal.get('Green Folder Meeting Date'), errors='coerce')

    if pd.notna(green_start) and pd.notna(green_end):
        gantt_data['Task'].append('Green Folder Dates')
        gantt_data['Start'].append(green_start)
        gantt_data['Finish'].append(green_end)
        gantt_data['Status'].append('Green Folder')
        gantt_data['Color'].append('purple')

    # Adding task data with conditional coloring based on status and dates
    for _, task in filtered_tasks_df.iterrows():
        start_date = pd.to_datetime(task['Start Date'], errors='coerce') or current_date
        due_date = pd.to_datetime(task['Due Date'], errors='coerce') or start_date + timedelta(days=1)
        status = task['Status Reason']

        gantt_data['Task'].append(task['Subject'])
        gantt_data['Start'].append(start_date)
        gantt_data['Finish'].append(due_date)
        gantt_data['Status'].append(status)

        # Apply color based on status and proximity to the due date
        if status == 'Completed':
            actual_end = pd.to_datetime(task.get('Actual End', due_date), errors='coerce')
            gantt_data['Finish'][-1] = actual_end
            gantt_data['Color'].append('gray')
        elif status == 'In Progress':
            if current_date > due_date:
                gantt_data['Color'].append('red')  # Past due (Overdue)
            elif current_date <= due_date <= current_date + timedelta(days=5):
                gantt_data['Color'].append('orange')  # Due in 5 days
            elif current_date + timedelta(days=5) < due_date <= current_date + timedelta(days=15):
                gantt_data['Color'].append('yellow')  # Due in 15 days
            else:
                gantt_data['Color'].append('green')  # In progress
        else:
            gantt_data['Color'].append('blue')  # Default for other statuses

    gantt_df = pd.DataFrame(gantt_data)

    if not gantt_df.empty:
        # Move Contract Dates and Green Folder Dates to the top
        fig = px.timeline(
            gantt_df,
            x_start="Start",
            x_end="Finish",
            y="Task",
            title=f"Gantt Chart for {deal_name}",
            color="Color",
            category_orders={'Task': ['Contract Dates', 'Green Folder Dates', 'Actual Contract Execution Date'] + gantt_df['Task'].tolist()},  # Ensure ordering
            color_discrete_map={
                'gray': 'gray',
                'green': 'green',
                'yellow': 'yellow',
                'orange': 'orange',
                'red': 'red',
                'teal': 'teal',
                'purple': 'purple',
                'blue': 'blue'
            },
            height=800
        )

        # Ensure ordering for the y-axis
        fig.update_yaxes(categoryorder="array", categoryarray=["Contract Dates", "Green Folder Dates", "Actual Contract Execution Date"])

        fig.update_layout(
            xaxis=dict(
                range=[current_date - timedelta(days=30), gantt_df['Finish'].max() + timedelta(days=30)],
                tickformat="%m/%d/%Y"
            ),
            showlegend=True
        )

        # Add vertical line (hash mark) for IP Expiration Date
        if pd.notna(ip_expiration):
            fig.add_vline(x=ip_expiration, line_dash="dash", line_color="black")

        return fig
    else:
        st.warning(f"No valid data to display in the Gantt chart for {deal_name}.")
        return None




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

            # Nice to show the cleaned columns. Debug statement written at top of screen. 
            # Debugging: Print the cleaned columns to verify "Actual Contract Execution Date" exists
            # st.write("Cleaned columns in deals_tasks_df:", deals_tasks_df.columns)

            # Clean 'Description' field in appointments
            if 'Description' in appointments_df.columns:
                appointments_df['Description'] = appointments_df['Description'].apply(strip_html)

            # Drop unwanted columns from appointments
            appointments_df = appointments_df.drop(columns=['Appointment', 'Row Checksum', '(Do Not Modify) Modified On'], errors='ignore')

            # Define expected columns (after cleaning)
            deal_columns = [
                'Regarding', 'Sub-Market', 'Calculated Deal Stage',
                'GF Submittal Date', 'Green Folder Meeting Date',
                'IP Expiration Date', 'Days to IP Expiration',
                'Projected Deal First Closing Date', 'Deal Homesite Total',
                'Homesite Size Description', 'Acquisition Type',
                'Primary Seller Company', 'Product Type Description',
                'CIC Final Approval Date', 'Actual Contract Execution Date'
            ]

            task_columns = [
                'Subject', 'Owner', 'Start Date', 'Due Date', 'Actual End',
                'Status Reason', 'Vendor Assigned', 'Task Category', 
                'Modified On', 'Comment'  
            ]

            # Extract deals data
            deals_df = deals_tasks_df[deal_columns].drop_duplicates().reset_index(drop=True)

            # Debug statement for sample data - written at top of screen
            # Ensure "Actual Contract Execution Date" is loaded
            # st.write("Sample of deals_df (checking for 'Actual Contract Execution Date'):", deals_df.head())

            # Handle non-finite values in 'Days to IP Expiration'
            deals_df['Days to IP Expiration'] = deals_df['Days to IP Expiration'].fillna(0).round().astype(int)
            
            # Convert date fields to datetime and extract only the date
            date_columns = ['GF Submittal Date', 'Green Folder Meeting Date', 'IP Expiration Date', 'Projected Deal First Closing Date', 'CIC Final Approval Date', 'Actual Contract Execution Date']
            #deals_df[date_columns] = deals_df[date_columns].apply(lambda x: pd.to_datetime(x, errors='coerce').dt.strftime('%m/%d/%Y'))
            # Convert all date columns and remove time component
            for col in date_columns:
                deals_df[col] = pd.to_datetime(deals_df[col], errors='coerce').dt.strftime('%m/%d/%Y')

            # Other code for tasks, appointments, and filtering...

            appointment_columns = [
                'Subject', 'Regarding', 'Owner', 'Status', 'Start Time',
                'End Time', 'Category', 'Description' 
            ]

            # Extract deals data
            deals_df = deals_tasks_df[deal_columns].drop_duplicates().reset_index(drop=True)

            # Handle non-finite values in 'Days to IP Expiration'
            deals_df['Days to IP Expiration'] = deals_df['Days to IP Expiration'].fillna(0).round().astype(int)
            
            # Convert date fields to datetime and extract only the date
            date_columns = ['GF Submittal Date', 'Green Folder Meeting Date', 'IP Expiration Date', 'Projected Deal First Closing Date', 'CIC Final Approval Date', 'Actual Contract Execution Date']
            deals_df[date_columns] = deals_df[date_columns].apply(lambda x: pd.to_datetime(x, errors='coerce').dt.strftime('%m/%d/%Y'))

            # Extract tasks data
            tasks_df = deals_tasks_df[task_columns + ['Regarding']].drop_duplicates().reset_index(drop=True)

            # Ensure 'Actual End' exists
            if 'Actual End' not in tasks_df.columns:
                tasks_df['Actual End'] = pd.NaT  # Ensure the column exists to avoid errors

            # Ensure unique columns before further processing
            if tasks_df.columns.duplicated().any():
                tasks_df = tasks_df.loc[:, ~tasks_df.columns.duplicated()]

            # Format date fields, including "Actual End", "Start Date", "Due Date", and "Modified On"
            date_columns = ['Start Date', 'Due Date', 'Modified On', 'Actual End']
            existing_date_columns = [col for col in date_columns if col in tasks_df.columns]

            # Convert to datetime
            if existing_date_columns:
                tasks_df[existing_date_columns] = tasks_df[existing_date_columns].apply(lambda x: pd.to_datetime(x, errors='coerce'))

            # Format the dates as mm/dd/yyyy
            tasks_df[existing_date_columns] = tasks_df[existing_date_columns].apply(lambda x: x.dt.strftime('%m/%d/%Y'))

            # You can then sort the DataFrame by any date column as needed
            tasks_df = tasks_df.sort_values(by='Actual End', ascending=True)

            # Extract and clean appointments data
            appointments_df[['Start Time', 'End Time']] = appointments_df[['Start Time', 'End Time']].apply(lambda x: pd.to_datetime(x, errors='coerce'))

            # Convert 'Modified On' to datetime if it exists
            if 'Modified On' in appointments_df.columns:
                appointments_df['Modified On'] = pd.to_datetime(appointments_df['Modified On'], errors='coerce')

            # Format the dates as mm/dd/yyyy
            appointments_df[['Modified On', 'Start Time', 'End Time']] = appointments_df[['Modified On', 'Start Time', 'End Time']].apply(lambda x: x.dt.strftime('%m/%d/%Y'))

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

            # Adjust columns to decrease space between buttons by using narrower column ratios
            col1, col2, col3, col4, col5, col6, col7, col8 = st.columns([0.1, 2.5, 1.8, 1.5, 1.1, 1.5, 1.5, 1.5])

            with col2:
                if st.button(f"Greenfolder Approved, Not Yet Closed ({len(greenfolder_approved_df)})"):
                    st.session_state['deal_filter'] = 'Greenfolder Approved, Not Yet Closed'
                    st.session_state['filtered_deals_df'] = greenfolder_approved_df
            with col3:
                if st.button(f"Green Folder Schedule ({len(green_folder_schedule_df)})"):
                    st.session_state['deal_filter'] = 'Green Folder Schedule'
                    st.session_state['filtered_deals_df'] = green_folder_schedule_df
            with col4:
                if st.button(f"Letters of Intent ({len(letters_of_intent_df)})"):
                    st.session_state['deal_filter'] = 'Letters of Intent'
                    st.session_state['filtered_deals_df'] = letters_of_intent_df 
            with col5:
                total_deals_count = len(deals_df)
                if st.button(f"All Deals ({total_deals_count})"):
                    st.session_state.pop('filtered_deals_df', None)  # Remove the filtered deals to reset to all deals
                    st.session_state.pop('deal_filter', None)  # Clear the deal filter state as well

            # Default to showing all deals if no button is clicked
            if 'filtered_deals_df' not in st.session_state:
                st.session_state['filtered_deals_df'] = deals_df

            filtered_deals_df = st.session_state['filtered_deals_df']

            # Convert date fields to datetime for sorting purposes
            date_columns = ['GF Submittal Date', 'Green Folder Meeting Date', 'IP Expiration Date', 'Projected Deal First Closing Date', 'CIC Final Approval Date']

            # Apply pd.to_datetime to each column individually, handling errors
            for col in date_columns:
                filtered_deals_df[col] = pd.to_datetime(filtered_deals_df[col], errors='coerce')

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

                # Sort by datetime fields properly before formatting them as strings
                if sort_column in date_columns:
                    filtered_deals_df = filtered_deals_df.sort_values(by=sort_column, ascending=(sort_order == "Ascending"))
                else:
                    filtered_deals_df = filtered_deals_df.sort_values(by=sort_column, ascending=(sort_order == "Ascending"))

            # After sorting, format the date columns for display
            # Only apply formatting to columns that are datetime
            for col in date_columns:
                if pd.api.types.is_datetime64_any_dtype(filtered_deals_df[col]):
                    filtered_deals_df[col] = filtered_deals_df[col].dt.strftime('%m/%d/%Y')

            # Now you can continue with other operations


            # Search Functionality using Dropdown with Search
            with st.expander("Search for Specific Deal"):
                deal_names = [""] + filtered_deals_df['Regarding'].dropna().unique().tolist()

                selected_deal = st.selectbox("Select a Deal:", deal_names)
                
                # Filter the DataFrame based on the selected deal
                if selected_deal:
                    filtered_deals_df = filtered_deals_df[filtered_deals_df['Regarding'] == selected_deal]

            # Add buttons to minimize/maximize all tasks and appointments
            col1, col2, col3, col4, col5, col6, col7, col8 = st.columns([.1, 1, 1, 1.5, 1.5, 1.5, 1.5, 1.5])

            with col2:
                minimize_all_button = st.button("Minimize All")
            with col3:
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
                go_to_download_link = f"<a href='#download' style='text-decoration: none; color: #015CAB;'>📥</a>"
                st.markdown(f"<h3>{deal_name} {return_to_top_link} {go_to_download_link}</h3>", unsafe_allow_html=True)

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
                ccol1, col2, col3, col4, col5, col6, col7, col8 = st.columns([0.1, 1.2, 1, 1, 1, 2, 2, 2])

                with col2:
                    if st.button(f"Show All Tasks ({total_tasks_count})", key=f"{deal_name}_show_all_{idx}"):
                        task_filter = "Show All"
                with col3:
                    if st.button(f"In Progress ({in_progress_count})", key=f"{deal_name}_in_progress_{idx}"):
                        task_filter = "In Progress"
                with col4:
                    if st.button(f"Completed ({completed_count})", key=f"{deal_name}_completed_{idx}"):
                        task_filter = "Completed"
                with col5:
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
        st.markdown('<a name="download"></a>', unsafe_allow_html=True)  # Anchor for download
        st.download_button(
            label="Download Excel",
            data=buffer.getvalue(),
            file_name=f"deal_task_appointment_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        # Excel icon URL (You can replace this URL with your own Excel icon)
        excel_icon_url = "https://storage.googleapis.com/absolute_gis_public/Images/lennar%20dashboard%20title.jpg"
        # Adding Excel icon next to Download button and rendering the button
        st.markdown(
            f"""
            <br><br><div style="display: flex; align-items: center;">
                <img src="{excel_icon_url}" width="500" style="margin-right: 10px;" />
            </div>
            """,
            unsafe_allow_html=True
        )
        st.markdown("<hr style='border: 4px solid #000;'>", unsafe_allow_html=True)

    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("Please upload exactly two Excel files: one for Deals/Tasks and one for Appointments.")