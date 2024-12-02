import streamlit as st


# Main page with guiding info (help section)
def main_page():
    st.title("Welcome to the Automated RFP Tool")
    st.markdown("#### :red[Placeholder text for quickstart guide here, for example:]")
    st.write(
        """
        

        This application provides the following tools:
        - **RFP Config:** Configure event-related parameters.
        - **Consolidate:** Aggregate and analyze RFP responses.
        
        Use the navigation menu on the left to select a tool.
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
