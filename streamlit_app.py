import streamlit as st
from script import script # Import your existing report generation function
from datetime import datetime
import io

# --- Page Configuration (makes your app look professional) ---
st.set_page_config(
    page_title="FIM Report Generator",
    page_icon="🚀",
    layout="centered"
)

# --- App UI ---
st.title("FIM Report Generator 🚀")
st.markdown("This tool automates the creation of a FIM report. Please provide the required details and upload the five necessary Excel files.")

# Use a form to prevent the app from re-running on every widget interaction
with st.form("report_form"):
    st.header("1. Report Details")
    # Use columns for a cleaner layout
    col1, col2 = st.columns(2)
    with col1:
        school_name = st.text_input("School Name", placeholder="e.g., Amity International School")
    with col2:
        principal_name = st.text_input("Principal's Name", placeholder="e.g., Dr. Jane Doe")

    st.header("2. Upload Excel Files")
    # The file_uploader returns an in-memory file object that pandas can read directly
    school_summary_file = st.file_uploader("School Summary File", type=["xlsx"])
    goals_data_file = st.file_uploader("Goals by Team File", type=["xlsx"])
    player_data_file = st.file_uploader("Player Data File", type=["xlsx"])
    educator_data_file = st.file_uploader("Educator Data File", type=["xlsx"])
    family_data_file = st.file_uploader("Family Data File", type=["xlsx"])

    # The submit button for the form
    submitted = st.form_submit_button("Generate Report")

# --- Backend Logic (runs only when the form is submitted) ---
if submitted:
    # First, validate that all inputs and files are present
    if not school_name or not principal_name:
        st.error("⚠️ Please fill in the School and Principal's Name.")
    elif not all([school_summary_file, goals_data_file, player_data_file, educator_data_file, family_data_file]):
        st.error("⚠️ Please upload all five required Excel files.")
    else:
        # If everything is ready, process the report
        try:
            # Prepare the data in the format your script expects


            # Show a spinner to let the user know something is happening
            with st.spinner("Analyzing data and building your report... This may take a moment."):
                # Call your existing function from script.py
                report_buffer = script(school_name,principal_name,goals_data_file,player_data_file,school_summary_file,educator_data_file,family_data_file)

            st.success("✅ Report generated successfully!")

            # Provide the download button for the generated file
            st.download_button(
                label="📥 Download Report (.docx)",
                data=report_buffer,
                file_name=f"{school_name.replace(' ', '_')}_FIM_Report_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"An error occurred during report generation: {e}")
            st.error("Please check that you have uploaded the correct files in the correct places.")