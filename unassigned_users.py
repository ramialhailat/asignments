import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Unassigned Users Mapping", layout="wide")
st.title("Unassigned Users Mapping")

# Sidebar: file uploads

# Sidebar menu with icons and navigation
st.sidebar.markdown("## 🗂️ Menu")
st.sidebar.markdown("- [🏠 Main Dashboard](./app.py)")
st.sidebar.markdown("- [👥 Unassigned Users Mapping](./unassigned_users.py)")
st.sidebar.markdown("---")
st.sidebar.markdown("### 📁 Upload Data Files")
project_file = st.sidebar.file_uploader("Upload Project Excel file", type=["xlsx", "xls"])
team_file = st.sidebar.file_uploader("Upload Team Excel file", type=["xlsx", "xls"])

@st.cache_data
def load_project(file):
    df = pd.read_excel(file)
    df['Parent'] = df['Parent'].fillna("")
    return df

@st.cache_data
def load_team(file):
    return pd.read_excel(file)

if project_file and team_file:
    df = load_project(project_file)
    team_df = load_team(team_file)

    # Prepare team members
    team_members = team_df['Name'].dropna().unique().tolist()

    # Count distinct epics per user
    project_counts = (
        df[df['EpicTitle'].notna()]
          .groupby('Assignee')['EpicTitle']
          .nunique()
          .reset_index(name='Number of Projects')
    ) if 'EpicTitle' in df.columns else pd.DataFrame({'Assignee':[], 'Number of Projects':[]})
    free_users = [u for u in team_members if u not in project_counts['Assignee'].tolist()]
    free_df = pd.DataFrame({'Assignee':free_users, 'Number of Projects':0})
    user_load = pd.concat([project_counts, free_df], ignore_index=True)

    st.subheader("Unassigned Users Mapping")
    updated_free = []
    for name in free_users:
        choices = ['--Keep--','Remove'] + user_load['Assignee'].tolist()
        mapping = st.selectbox(f"Map '{name}'", choices, key=name)
        if mapping == '--Keep--':
            updated_free.append(name)
        elif mapping == 'Remove':
            continue
        else:
            updated_free.append(mapping)

    st.subheader("Updated Unassigned Users")
    df_out = pd.DataFrame({'Name': updated_free})
    st.dataframe(df_out)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df_out.to_excel(writer, index=False)
    buf.seek(0)
    st.download_button(
        "Download Unassigned Users",
        buf,
        file_name='unassigned_users.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
else:
    st.info("Please upload both project and team files.")
