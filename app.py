import io
import pandas as pd
import plotly.express as px
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
import streamlit as st

st.set_page_config(
    page_title="Meeting Analysis App",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📊 Meeting Analysis App")
st.markdown("Management and tracking application for sales meeting frequency by project stage")

STAGES = [
    "Chance",
    "Get Design info",
    "Design In",
    "1st quotation",
    "Approval Sample",
    "Final Quotation",
    "Reliability Approval",
    "Design Win",
    "PO",
]
STAGE_COLS_DIDW = [
    "Chance.1",
    "Get Design info.1",
    "Design In.1",
    "1st quotation.1",
    "Approval Sample.1",
    "Final Quotation.1",
    "Reliability Approval.1",
    "Design Win.1",
    "PO.1",
]


@st.cache_data
def process_data(file_act_path, file_didw_path):
  df_act = pd.read_excel(file_act_path)
  df_didw = pd.read_excel(file_didw_path)

  df_act.columns = [c.strip(".").strip() for c in df_act.columns]
  df_didw.columns = [c.strip(".").strip() for c in df_didw.columns]

  meeting_types = ["meal", "meeting", "meeting(on-line)", "meeting online"]
  df_act_clean = df_act[
      df_act["Type"].astype(str).str.lower().isin(meeting_types)
  ].copy()
  df_act_clean["Date"] = pd.to_datetime(
      df_act_clean["Date"], errors="coerce"
  )

  for col in STAGE_COLS_DIDW:
    if col in df_didw.columns:
      df_didw[col] = pd.to_datetime(df_didw[col], errors="coerce")

  assigned_stages = []

  for _, act_row in df_act_clean.iterrows():
    sales = act_row.get("Sales Member")
    cust = act_row.get("Customer")
    pjt = act_row.get("Project Name")
    act_date = act_row["Date"]

    matched_didw = pd.DataFrame()
    if pd.notna(pjt):
      matched_didw = df_didw[
          (df_didw["Sales Member"] == sales)
          & (df_didw["Project Name"] == pjt)
      ]
    if matched_didw.empty and pd.notna(cust):
      matched_didw = df_didw[
          (df_didw["Sales Member"] == sales) & (df_didw["Customer Name"] == cust)
      ]

    assigned_stage = "Non-Relevant"

    if not matched_didw.empty and pd.notna(act_date):
      for _, didw_row in matched_didw.iterrows():
        dates = []
        for name, col in zip(STAGES, STAGE_COLS_DIDW):
          d = didw_row.get(col)
          if pd.notna(d):
            dates.append((name, d))

        if not dates:
          continue

        found = False
        for i in range(len(dates)):
          s_name, s_date = dates[i]
          next_date = dates[i + 1][1] if i + 1 < len(dates) else None

          if s_date <= act_date:
            if next_date is None or act_date < next_date:
              assigned_stage = s_name
              found = True
              break
        if found:
          break

    assigned_stages.append(assigned_stage)

  df_act_clean["Assigned_Stage"] = assigned_stages
  return df_act_clean, df_didw


try:
  df_processed, df_didw_raw = process_data(
      "20261029_Activity Detail.xlsx", "DIDW_12T (38).xlsx"
  )

  st.sidebar.header("🔍 Filters")

  sales_list = ["All"] + sorted(
      df_processed["Sales Member"].dropna().unique().tolist()
  )
  selected_sales = st.sidebar.selectbox("Sales Member", sales_list)

  df_temp_region = df_processed.copy()
  if selected_sales != "All":
    df_temp_region = df_temp_region[
        df_temp_region["Sales Member"] == selected_sales
    ]
  region_list = ["All"] + sorted(
      df_temp_region["Region"].dropna().unique().tolist()
  )
  selected_region = st.sidebar.selectbox("Region", region_list)

  df_temp_pjt = df_temp_region.copy()
  if selected_region != "All":
    df_temp_pjt = df_temp_pjt[df_temp_pjt["Region"] == selected_region]
  pjt_list = ["All"] + sorted(
      df_temp_pjt["Project Name"].dropna().unique().tolist()
  )
  selected_pjt = st.sidebar.selectbox("Project Name", pjt_list)

  min_date = df_processed["Date"].min().date()
  max_date = df_processed["Date"].max().date()

  st.sidebar.markdown("---")
  st.sidebar.markdown("📅 **Period**")
  start_date = st.sidebar.date_input("Start Date", value=min_date)
  end_date = st.sidebar.date_input("End Date", value=max_date)

  if "applied" not in st.session_state:
    st.session_state.applied = True

  if st.sidebar.button("✅ OK - Apply Report", type="primary"):
    st.session_state.applied = True

  df_filtered = df_processed.copy()
  if selected_sales != "All":
    df_filtered = df_filtered[df_filtered["Sales Member"] == selected_sales]
  if selected_region != "All":
    df_filtered = df_filtered[df_filtered["Region"] == selected_region]
  if selected_pjt != "All":
    df_filtered = df_filtered[df_filtered["Project Name"] == selected_pjt]

  if start_date and end_date:
    df_filtered = df_filtered[
        (df_filtered["Date"].dt.date >= start_date)
        & (df_filtered["Date"].dt.date <= end_date)
    ]

  # --- MAIN DASHBOARD ---
  col1, col2, col3 = st.columns(3)
  col1.metric("Total Filtered Meetings", len(df_filtered))
  col2.metric("Participating Sales", df_filtered["Sales Member"].nunique())
  col3.metric("Total Projects", df_filtered["Project Name"].nunique())

  st.markdown("---")
  st.subheader("📈 Meeting Frequency by Project Stage")

  categories_order = STAGES + ["Non-Relevant"]
  stage_counts = (
      df_filtered["Assigned_Stage"]
      .value_counts()
      .reindex(categories_order, fill_value=0)
      .reset_index()
  )
  stage_counts.columns = ["Stage", "Meeting Count"]

  fig = px.bar(
      stage_counts,
      x="Stage",
      y="Meeting Count",
      text="Meeting Count",
      color="Stage",
      color_discrete_map={"Non-Relevant": "#EF553B"},
  )
  fig.update_traces(textposition="outside")
  fig.update_layout(
      xaxis_title="Project Stage",
      yaxis_title="Meeting Count",
      showlegend=False,
  )
  st.plotly_chart(fig, use_container_width=True)

  st.subheader("📋 Detailed Data Table")
  st.dataframe(df_filtered, use_container_width=True)

  # --- EXCEL EXPORT ---
  wb = openpyxl.Workbook()

  ws_summary = wb.active
  ws_summary.title = "Summary & Chart"

  for r in dataframe_to_rows(stage_counts, index=False, header=True):
    ws_summary.append(r)

  chart = BarChart()
  chart.type = "col"
  chart.style = 10
  chart.title = "Sales Meeting Frequency by Project Stage"
  chart.y_axis.title = "Meeting Count"
  chart.x_axis.title = "Project Stage"

  data_ref = Reference(
      ws_summary, min_col=2, min_row=1, max_row=len(stage_counts) + 1
  )
  cats_ref = Reference(
      ws_summary, min_col=1, min_row=2, max_row=len(stage_counts) + 1
  )

  chart.add_data(data_ref, titles_from_data=True)
  chart.set_categories(cats_ref)
  chart.legend = None

  chart.height = 12
  chart.width = 20
  ws_summary.add_chart(chart, "D2")

  ws_details = wb.create_sheet(title="Filtered Meeting Details")
  df_details_export = df_filtered.copy()
  if "Date" in df_details_export.columns:
    df_details_export["Date"] = df_details_export["Date"].astype(str)

  for r in dataframe_to_rows(df_details_export, index=False, header=True):
    ws_details.append(r)

  ws_raw = wb.create_sheet(title="Original Raw Data")
  df_raw_export = pd.read_excel("20261029_Activity Detail.xlsx")
  if "Date" in df_raw_export.columns:
    df_raw_export["Date"] = df_raw_export["Date"].astype(str)

  for r in dataframe_to_rows(df_raw_export, index=False, header=True):
    ws_raw.append(r)

  output = io.BytesIO()
  wb.save(output)
  excel_data = output.getvalue()

  st.download_button(
      label="📥 Download Complete Report (.xlsx)",
      data=excel_data,
      file_name="Meeting_Analysis_Report.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  )

except Exception as e:
  st.error(f"Error reading data files: {e}")
