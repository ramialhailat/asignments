import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime, timedelta
import os

# Attempt to import AgGrid for interactive tables
try:
    from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
    USE_AGGRID = True
except ImportError:
    USE_AGGRID = False

st.set_page_config(page_title="Epic Assignment Dashboard", layout="wide")
st.title("Epic Assignment Dashboard")
hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;}
header {visibility: visible;}
footer {visibility: hidden;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)
# Sidebar: file uploads
st.sidebar.header("Upload Data Files")

remove_cache = st.sidebar.button("Remove Cached Files")
project_file = st.sidebar.file_uploader("Upload Project Excel file", type=["xlsx", "xls"])
team_file = st.sidebar.file_uploader("Upload Team Excel file", type=["xlsx", "xls"])

# Remove cached files if button pressed
project_path = "project_upload.xlsx"
team_path = "team_upload.xlsx"
if remove_cache:
    if os.path.exists(project_path):
        os.remove(project_path)
    if os.path.exists(team_path):
        os.remove(team_path)
    st.sidebar.success("Cached files removed. Please upload new files.")

# Paths for persisted uploads
project_path = "project_upload.xlsx"
team_path = "team_upload.xlsx"

# Save uploaded files to disk
if project_file:
    with open(project_path, "wb") as f:
        f.write(project_file.getbuffer())
if team_file:
    with open(team_path, "wb") as f:
        f.write(team_file.getbuffer())

@st.cache_data
def load_project(file):
    df = pd.read_excel(file)
    df['Parent'] = df['Parent'].fillna("")
    return df

@st.cache_data
def load_team(file):
    return pd.read_excel(file)

if os.path.exists(project_path) and os.path.exists(team_path):
    # Load data from persisted files
    df = load_project(project_path)
    team_df = load_team(team_path)

    # Required project columns
    required_cols = ['Hierarchy','Title','Parent','Assignee',
                     'Project Actual Start Date','Due date']
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Missing required columns: {missing}")
        st.stop()

    # Resolve EpicTitle for each row
    hierarchy_map = df.set_index('Title')['Hierarchy'].to_dict()
    parent_map = df.set_index('Title')['Parent'].to_dict()
    def find_epic(item):
        current = item
        while current:
            if hierarchy_map.get(current) == 'Epic':
                return current
            current = parent_map.get(current, '')
        return None

    df['EpicTitle'] = df.apply(
        lambda r: r['Title'] if r['Hierarchy']=='Epic' else find_epic(r['Parent']),
        axis=1
    )

    # Prepare team members
    team_members = team_df['Name'].dropna().unique().tolist()

    # Count distinct epics per user
    project_counts = (
        df[df['EpicTitle'].notna()]
          .groupby('Assignee')['EpicTitle']
          .nunique()
          .reset_index(name='Number of Projects')
    )
    # Add free users
    free_users = [u for u in team_members if u not in project_counts['Assignee'].tolist()]
    free_df = pd.DataFrame({'Assignee':free_users, 'Number of Projects':0})
    user_load = pd.concat([project_counts, free_df], ignore_index=True)

    # Compute free date
    def get_free_date(user):
        epics = df[df['Assignee']==user]['EpicTitle'].dropna().unique()
        if len(epics)==0:
            return pd.to_datetime(datetime.today())
        dates = []
        for e in epics:
            epic_row = df[(df['Hierarchy']=='Epic') & (df['Title']==e)]
            if not epic_row.empty:
                dates.append(epic_row['Due date'].iloc[0])
        return pd.to_datetime(max(dates)) if dates else pd.to_datetime(datetime.today())

    user_load['Free_Date'] = user_load['Assignee'].apply(get_free_date)

    # Build weekly availability with emojis, extended to end of year
    today = datetime.today()
    week0 = today - timedelta(days=today.weekday())
    # Find last week of the year
    last_day = datetime(today.year, 12, 31)
    last_week = last_day - timedelta(days=last_day.weekday())
    weeks = []
    w = week0
    while w <= last_week:
        weeks.append(w)
        w += timedelta(weeks=1)
    week_cols = []
    for w in weeks:
        col = w.strftime('%Y-%m-%d')
        week_cols.append(col)
        user_load[col] = user_load['Free_Date'].apply(
            lambda d: '🟢' if d <= pd.to_datetime(w) else '🔴'
        )

    # Tabs for dashboard and unassigned users mapping
    tab1, tab2 = st.tabs(["Dashboard", "Unassigned Users Mapping"])

    with tab1:
        st.subheader("Projects Assigned per User")
        fig = px.bar(
            project_counts,
            x='Assignee', y='Number of Projects',
            title='Number of Projects per User',
            labels={'Number of Projects':'# Projects'}
        )
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("User Load & Availability")
        if USE_AGGRID:
            # Prepare exportable DataFrame with 'Free'/'Busy' instead of emojis
            export_df = user_load.drop(columns=['Free_Date']).copy()
            for col in week_cols:
                export_df[col] = export_df[col].apply(lambda v: 'Free' if v == '🟢' else ('Busy' if v == '🔴' else v))
            display_df = user_load.drop(columns=['Free_Date'])
            gb = GridOptionsBuilder.from_dataframe(display_df)
            gb.configure_selection('single')
            grid_opts = gb.build()
            grid_response = AgGrid(
                display_df,
                gridOptions=grid_opts,
                update_mode=GridUpdateMode.SELECTION_CHANGED,
                fit_columns_on_grid_load=True
            )
            selected_rows = grid_response.get('selected_rows', [])
            selected_user =None
            if isinstance(selected_rows, list) and len(selected_rows) > 0:
                selected_user = selected_rows[0].get('Assignee')
            elif hasattr(selected_rows, 'empty') and hasattr(selected_rows, 'iloc') and not selected_rows.empty:
                selected_user = selected_rows.iloc[0]['Assignee']
            # Export button for User Load & Availability table
            buf_user_load = io.BytesIO()
            with pd.ExcelWriter(buf_user_load, engine='openpyxl') as writer:
                export_df.to_excel(writer, index=False)
            buf_user_load.seek(0)
            st.download_button(
                "Download User Load & Availability",
                buf_user_load,
                file_name='user_load_availability.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            st.warning("Install `streamlit-aggrid` for interactive table selection.")
            export_df = user_load.drop(columns=['Free_Date']).copy()
            for col in week_cols:
                export_df[col] = export_df[col].apply(lambda v: 'Free' if v == '🟢' else ('Busy' if v == '🔴' else v))
            display_df = user_load.drop(columns=['Free_Date'])
            st.dataframe(display_df)
            selected_user = st.selectbox("Select user for details", user_load['Assignee'])
            # Export button for User Load & Availability table
            buf_user_load = io.BytesIO()
            with pd.ExcelWriter(buf_user_load, engine='openpyxl') as writer:
                export_df.to_excel(writer, index=False)
            buf_user_load.seek(0)
            st.download_button(
                "Download User Load & Availability",
                buf_user_load,
                file_name='user_load_availability.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        if selected_user:
            st.markdown(f"### Assignments for **{selected_user}**")
            detail_cols = ['Hierarchy','Title','EpicTitle','Project Actual Start Date','Due date']
            if 'Work item key' in df.columns:
                detail_cols.insert(0,'Work item key')
            elif 'Key' in df.columns:
                detail_cols.insert(0,'Key')

            selected_df = df[df['Assignee']==selected_user][detail_cols]
            # If Work item key exists, add a link column
            if 'Work item key' in selected_df.columns:
                # Example: link to Jira, replace URL as needed
                base_url = "https://networkinternational.atlassian.net/browse/"
                selected_df['Work item link'] = selected_df['Work item key'].apply(
                    lambda k: f"[{k}]({base_url}{k})" if pd.notnull(k) else ""
                )
                # Move link column to front
                cols_with_link = ['Work item link'] + [c for c in selected_df.columns if c != 'Work item link']
                st.markdown(selected_df.to_markdown(index=False), unsafe_allow_html=True)
            else:
                st.dataframe(selected_df)

            free_date = user_load.loc[user_load['Assignee']==selected_user, 'Free_Date'].iloc[0]
            st.markdown(f"**Estimated Free Date:** {free_date.date()}")

    with tab2:
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
