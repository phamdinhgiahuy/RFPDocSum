import streamlit as st

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False


def login():
    if st.button("Log in"):
        st.session_state.logged_in = True
        st.rerun()


def logout():
    if st.button("Log out"):
        st.session_state.logged_in = False
        st.rerun()


login_page = st.Page(login, title="Log in", icon=":material/login:")
logout_page = st.Page(logout, title="Log out", icon=":material/logout:")

dashboard = st.Page(
    "admin/dashboard.py", title="Dashboard", icon=":material/dashboard:", default=True
)
bugs = st.Page("admin/bugs.py", title="Bugs", icon=":material/bug_report:")
alerts = st.Page(
    "admin/alerts.py", title="System alerts", icon=":material/notification_important:"
)

config = st.Page(
    "tools/event_config.py", title="RFP Config", icon=":material/settings:"
)
consolidate = st.Page(
    "tools/consolidate.py", title="Consolidate", icon=":material/compare:"
)

if st.session_state.logged_in:
    pg = st.navigation(
        {
            "Account": [logout_page],
            "Admin": [dashboard, bugs, alerts],
            "Tools": [config, consolidate],
        }
    )
else:
    pg = st.navigation([login_page])

pg.run()
