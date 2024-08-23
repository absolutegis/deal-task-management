import streamlit as st
import pandas as pd

# Set the page title and layout
st.set_page_config(page_title="Deal and Task Management Interface", layout="wide")

# Title of the app
st.title("Deal and Task Management Interface")

# File uploader for the Excel file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Load the uploaded Excel file
    df = pd.read_excel(uploaded_file)

    # Step 1: Extract Unique Deals (Columns A to M)
    deal_columns = df.columns[:13]  # Assuming columns A to M are the first 13 columns
    deals_df = df[deal_columns].drop_duplicates()

    # Step 2: Relate Tasks (Columns N to V) to the Unique Identifier in Column A
    task_columns = df.columns[[0] + list(range(13, len(df.columns)))]  # Columns A (Identifier) + N to V
    tasks_df = df[task_columns]

    # List of expected task columns
    expected_task_columns = ['Regarding', 'Task Category', 'Owner', 'Subject', 'Start Date', 'Due Date', 'Vendor Assigned', 'Priority', 'Comment', 'Status Reason']

    # Loop through each deal and display its related tasks horizontally
    for _, deal in deals_df.iterrows():
        # Display the deal information horizontally (only A to M)
        st.subheader(f"Deal: {deal['Regarding']}")
        st.write(deal.to_frame().T)  # Display only columns A to M

        # Filter tasks related to this deal
        related_tasks = tasks_df[tasks_df['Regarding'] == deal['Regarding']]

        # Apply conditional formatting within the DataFrame
        if not related_tasks.empty:
            # Ensure the columns exist before displaying them
            available_columns = [col for col in expected_task_columns if col in related_tasks.columns]
            related_tasks = related_tasks[available_columns]

            # Apply basic formatting for overdue tasks and status
            def highlight_tasks(s):
                if s['Status Reason'] == 'Completed':
                    return ['background-color: gray; color: white'] * len(s)
                elif s['Status Reason'] == 'In Progress':
                    return ['background-color: green; color: white'] * len(s)
                elif not pd.isna(s['Due Date']) and pd.to_datetime(s['Due Date']) < pd.to_datetime('today'):
                    return ['background-color: red; color: white'] * len(s)
                return [''] * len(s)

            styled_tasks = related_tasks.style.apply(highlight_tasks, axis=1)

            # Collapsible section for related tasks with sorting enabled
            with st.expander("Related Tasks:"):
                st.dataframe(styled_tasks)  # Display the tasks with sorting enabled
else:
    st.info("Please upload an Excel file to begin.")
