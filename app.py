import streamlit as st

st.set_page_config(
    page_title="University Academic Analysis",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 University Academic Analysis")
st.subheader("Academic Performance & Attendance Dashboards")

st.markdown("""
### Select a dashboard from the left sidebar

- 📈 **Attendance Analysis**
- 📝 **MSE & ESE Analysis**

Use the sidebar to switch between the two dashboards. Both are available through this **single web link**.
""")

st.info("Upload the relevant Academia ERP report inside the selected dashboard.")
