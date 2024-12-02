import streamlit as st


# Main page with guiding info (help section)
def main_page():
    st.title("Welcome to the Automated RFP Tool")

    st.markdown("#### :red[Where to start?]")

    st.write(
        """
        This application provides the following tools:
        1. **RFP Config:** Configure RFP event-related settings for the consolidation file.
        2. **Consolidate:** Aggregate and analyze RFP responses from multiple vendors.
        
        👈 Use the navigation menu on the left to select a tool.
        Expand the help section for a more detailed guide.
        """
    )
    with st.expander("See Full Guide"):
        st.markdown(
            """
            ### Step 1: Choose File Configuration
            Decide if the questionnaire and pricing contents are in a single file or separate Excel files.
            
            ### Step 2: Upload Files
            **Navigate to the RFP Config section of the tool.**
            
            **Upload template files:**  
            - Move to the “Template Files section” and upload the template files (the files that were sent to the vendors).  
            - Select the template file that contains the questionnaire and pricing respectively.
            
            **Upload vendor files:**  
            - Move to the "Supplier Response Files" section to upload files received from each vendor.  
            - Ensure all files are in `.xlsx` format.  
            - The tool will validate the uploaded files.
            
            ### Step 3: Select Consolidation Method
            Choose between the following options:  
            - **Side-by-Side Consolidation:** Vendor responses are aligned in a single sheet.  
            - **Separate Sheet Consolidation:** Each vendor’s responses are placed in its own sheet.  
            
            If required, enable the Summarization option to generate concise summaries of responses.
            
            ### Step 4: Configure Settings
            Move to the consolidate section of the tool to specify which sheets to include in the consolidation process and generate consolidated files.
            
            ### Step 5: Consolidate Data and Download Results
            - Click the **Consolidate Data** button.  
            - Wait for the tool to process the files and combine the data.  
            - Click the **Download Consolidated File** button.  
            - Save the file locally for further analysis.
            """
        )


config = st.Page(
    "tools/event_config.py", title="RFP Config", icon=":material/settings:"
)
consolidate = st.Page(
    "tools/consolidate.py", title="Consolidate", icon=":material/compare:"
)

pg = st.navigation(
    {
        "Help": [st.Page(main_page, title="Help", icon=":material/help:")],
        "Tools": [config, consolidate],
    }
)
pg.run()
