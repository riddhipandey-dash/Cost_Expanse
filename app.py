import pandas as pd
import plotly.express as px
import streamlit as st
def load_and_clean(path: str) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=0, header=None)
    months = raw.iloc[1].ffill()
    sub = raw.iloc[2]
    headers = []
    for i, (m, s) in enumerate(zip(months, sub)):
        if i in (0, 1):
            headers.append(str(s))
        else:
            headers.append(f"{str(m).strip()}_{str(s).strip()}")
    df = raw.iloc[3:].copy()
    df.columns = headers
    df = df.rename(columns={"S.NO.": "S_NO", "STATE": "STATE"})
    df = df[df["STATE"].notna()]
    df["STATE"] = df["STATE"].astype(str).str.strip()
    df = df[~df["STATE"].str.upper().isin(["TOTAL", "NOC"])]
    value_cols = [c for c in df.columns if "_" in c]
    long_df = df.melt(id_vars=["S_NO", "STATE"], value_vars=value_cols, var_name="MONTH_METRIC", value_name="VALUE")
    month_metric = long_df["MONTH_METRIC"].str.split("_", n=1, expand=True)
    long_df["MONTH"] = month_metric[0].str.title()
    long_df["METRIC"] = month_metric[1].str.upper()
    long_df["VALUE"] = (
        long_df["VALUE"].astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
    )
    long_df["VALUE"] = pd.to_numeric(long_df["VALUE"], errors="coerce")
    pivot = long_df.pivot_table(
        index=["STATE", "MONTH"],
        columns="METRIC",
        values="VALUE",
        aggfunc="sum",
    ).reset_index()
    for col in ["SALE", "SALARY", "EXP"]:
        if col not in pivot.columns:
            pivot[col] = 0
    if "TOTAL" in pivot.columns:
        pivot = pivot.drop(columns=["TOTAL"])
    pivot["TOTAL_EXPENSE"] = pivot["SALARY"].fillna(0) + pivot["EXP"].fillna(0)
    return pivot

def month_order(df: pd.DataFrame) -> list:
    order = ["April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"]
    existing = [m for m in order if m in df["MONTH"].unique()]
    remaining = [m for m in df["MONTH"].unique() if m not in existing]
    return existing + remaining

DATA_PATH = "SALE & EXPANSES DATA 2025-2026.xlsx"
data = load_and_clean(DATA_PATH)

st.set_page_config(page_title="Sales & Expense Dashboard", layout="wide", page_icon="📊", initial_sidebar_state="expanded")
px.defaults.template = "plotly_white"
px.defaults.color_continuous_scale = px.colors.sequential.Teal

# --- Sidebar Filters ---
st.title("Sales & Expense Dashboard (FY 2025-2026)")
st.caption("Cleaned and visualized for non-technical readers")
all_states = sorted(data["STATE"].dropna().unique())
all_months = month_order(data)
state_options = ["All"] + all_states
month_options = ["All"] + all_months
selected_states = st.sidebar.multiselect("State", state_options, default=["All"], key="sidebar_state_select")
selected_months = st.sidebar.multiselect("Month", month_options, default=["All"], key="sidebar_month_select")
state_filter = all_states if "All" in selected_states or not selected_states else selected_states
month_filter = all_months if "All" in selected_months or not selected_months else selected_months
filtered = data[data["STATE"].isin(state_filter) & data["MONTH"].isin(month_filter)].copy()

# --- KPI Section ---
kpi1, kpi2 = st.columns(2)
kpi1.metric("Total Sales (₹)", f"₹ {filtered['SALE'].sum():,.0f}")
kpi2.metric("Total Expense (₹)", f"₹ {filtered['TOTAL_EXPENSE'].sum():,.0f}")
st.markdown("---")

# --- Sales & Expense by State ---
st.subheader("Sales & Expense by State (2025)")
state_agg = filtered.groupby("STATE", as_index=False)[["SALE", "TOTAL_EXPENSE"]].sum().sort_values("SALE", ascending=False)
fig_state_sales = px.bar(state_agg, x="STATE", y="SALE", color="SALE", color_continuous_scale="Blues", title="Sales by State (₹)")
fig_state_sales.update_layout(yaxis_tickprefix="₹ ", height=400)
fig_state_exp = px.bar(state_agg, x="STATE", y="TOTAL_EXPENSE", color="TOTAL_EXPENSE", color_continuous_scale="Oranges", title="Expense by State (₹)")
fig_state_exp.update_layout(yaxis_tickprefix="₹ ", height=400)
col1, col2 = st.columns(2)
col1.plotly_chart(fig_state_sales, use_container_width=True, key="state_sales_chart")
top_state = state_agg.iloc[0] if not state_agg.empty else None
bottom_state = state_agg.iloc[-1] if not state_agg.empty else None
col1.markdown(f"**Insight:** Highest sales in {top_state['STATE']} (₹ {top_state['SALE']:,.0f}), lowest in {bottom_state['STATE']} (₹ {bottom_state['SALE']:,.0f})." if top_state is not None and bottom_state is not None else "")
col2.plotly_chart(fig_state_exp, use_container_width=True, key="state_exp_chart")
top_exp_state = state_agg.sort_values('TOTAL_EXPENSE', ascending=False).iloc[0] if not state_agg.empty else None
bottom_exp_state = state_agg.sort_values('TOTAL_EXPENSE', ascending=True).iloc[0] if not state_agg.empty else None
col2.markdown(f"**Insight:** Highest expense in {top_exp_state['STATE']} (₹ {top_exp_state['TOTAL_EXPENSE']:,.0f}), lowest in {bottom_exp_state['STATE']} (₹ {bottom_exp_state['TOTAL_EXPENSE']:,.0f})." if top_exp_state is not None and bottom_exp_state is not None else "")
st.markdown("---")

# --- Monthly Trends ---
st.subheader("Monthly Sales & Expense Trends (2025)")
monthly = filtered.groupby("MONTH", as_index=False)[["SALE", "TOTAL_EXPENSE", "SALARY", "EXP"]].sum()
monthly["MONTH"] = pd.Categorical(monthly["MONTH"], categories=month_order(filtered), ordered=True)
monthly = monthly.sort_values("MONTH")
fig_monthly = px.line(monthly, x="MONTH", y=["SALE", "TOTAL_EXPENSE"], markers=True, title="Monthly Sales vs Expense", color_discrete_sequence=["#2c7fb8", "#f58518"])
fig_monthly.update_layout(yaxis_tickprefix="₹ ", height=400)
st.plotly_chart(fig_monthly, use_container_width=True, key="monthly_trend_chart")
best_month = monthly.sort_values("SALE", ascending=False).iloc[0] if not monthly.empty else None
worst_month = monthly.sort_values("SALE", ascending=True).iloc[0] if not monthly.empty else None
st.markdown(f"**Insight:** Peak sales in {best_month['MONTH']} (₹ {best_month['SALE']:,.0f}), lowest in {worst_month['MONTH']} (₹ {worst_month['SALE']:,.0f})." if best_month is not None and worst_month is not None else "")

# --- Salary vs Other Expense (Stacked Bar) ---
st.subheader("Monthly Salary vs Other Expense (2025)")
fig_salary_exp = px.bar(monthly, x="MONTH", y=["SALARY", "EXP"], title="Salary vs Other Expense (₹)", color_discrete_sequence=["#756bb1", "#8c6d31"])
fig_salary_exp.update_layout(barmode="stack", yaxis_tickprefix="₹ ", height=400)
st.plotly_chart(fig_salary_exp, use_container_width=True, key="salary_exp_chart")
max_salary_month = monthly.sort_values("SALARY", ascending=False).iloc[0] if not monthly.empty else None
max_exp_month = monthly.sort_values("EXP", ascending=False).iloc[0] if not monthly.empty else None
st.markdown(f"**Insight:** Highest salary payout in {max_salary_month['MONTH']} (₹ {max_salary_month['SALARY']:,.0f}), highest other expense in {max_exp_month['MONTH']} (₹ {max_exp_month['EXP']:,.0f})." if max_salary_month is not None and max_exp_month is not None else "")
st.markdown("---")

# --- Sales Heatmap (State vs Month) ---
st.subheader("Sales Heatmap (State vs Month)")
heatmap_data = filtered.pivot_table(index="STATE", columns="MONTH", values="SALE", aggfunc="sum", fill_value=0)
heatmap_data = heatmap_data.loc[state_agg["STATE"]]  # order by total sales
fig_heatmap = px.imshow(heatmap_data, labels=dict(x="Month", y="State", color="Sales (₹)"), aspect="auto", color_continuous_scale="Blues")
fig_heatmap.update_layout(height=500)
st.plotly_chart(fig_heatmap, use_container_width=True, key="heatmap_chart")
