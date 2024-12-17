import streamlit as st


# Main page with guiding info (help section)
def main_page():
    # st.image(r"", width=200)
    st.title("Welcome to the Automated RFP Tool")

    st.markdown("### :blue[▶️ Where to start?]")

    st.markdown(
        """
        ***Auto RFP Analyzer:** Simplifies consolidating and analyzing questionnaire and pricing data from vendors.*
        This application provides the following tools:

        1. ⚙️ **RFP Config:** 
        Configure RFP event-related settings for the consolidation file.
        2. 📋 **Consolidate:** 
        Aggregate and analyze RFP responses from multiple suppliers either side-by-side or in separate sheets.
        
        👈 Use the *navigation menu* on the left to select a tool.
        
        *Expand the ℹ️ section below for a more detailed guide.*
        """
    )
    with st.expander("ℹ️ Full User Guide"):
        st.markdown(
            """
        ### 📱 Features
        #### 💲Pricing
        - **Organized Data**: Differentiates descriptions and prices.  
        - **Color-Coded Pricing**: Highlights vendor-specific values.  
        - **Aggregation Options**:  
        - **Side-by-Side**: Prices from multiple vendors in one sheet.  
        - **Sheet-by-Sheet**: Vendor prices in separate sheets.  
        - **Analysis Tools**: Generates summaries and diagrams.

        #### ❔Questionnaire
        - **Upload Files**: Accepts valid `.xlsx` files.
        - **Parse Responses**: Matches template columns, highlights mismatched rows, and extracts vendor data.
        - **Consolidation Options**:  
        - **Side-by-Side**: All responses in one sheet.  
        - **Separate Sheets**: Each vendor's data in its own sheet.  
        - **Summarization**: Option to create concise summaries.

        ---

        ### 📖 Instructions

        #### 🚀 Launch
        Run in terminal:  
        ```bash
        streamlit run [file_path]
        ```
        Replace `[file_path]` with the app file's path.


        #### 🪜 Steps

        1. **Choose File Organization**: Single or separate questionnaire and pricing files.  
        2. **Upload Files**:  
        - **Templates**: Upload questionnaire and pricing templates.  
        - **Vendor Files**: Upload `.xlsx` responses.  
        3. **Select Consolidation**:  
        - **Side-by-Side**: All in one sheet.  
        - **Separate Sheets**: Vendor-specific sheets.  
        - Optionally enable **Summarization**.  
        4. **Configure Settings**: Select sheets to consolidate.  
        5. **Consolidate and Download**:  
        - Click **Consolidate Data**.  
        - After processing, click **Download Consolidated File**.


        #### ❗Error Handling

        - **Invalid File**: Upload a valid `.xlsx` file.  
        - **No Matching Columns**: Verify templates match vendor files.  
        - **Missing Summarization**: Check input for meaningful content.

        ---

        ### ☑️ Tips 

        - Align vendor files with template structure.  
        - Use **Side-by-Side Consolidation** for easy comparisons.  
        - Enable **Summarization** for quick reviews.  
        - Enhance pricing diagrams by adjusting chart elements (e.g., Axes, Data Labels).
        - **⚠️Warning messages** can be generated when opening the consolidated file in Excel especially when suppliers use Rich Text formats in their responses. This can be ignored, click **Yes** to proceed.
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
