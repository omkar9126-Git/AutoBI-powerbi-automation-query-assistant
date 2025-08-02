
import pandas as pd
import streamlit as st
import altair as alt

def get_autocomplete_options(df: pd.DataFrame):
    return {
        "heats": sorted(df["Heat No."].dropna().unique().tolist()),
        "materials": sorted(df["Material"].dropna().unique().tolist()),
        "sections": sorted(df["Section"].dropna().unique().tolist())
    }

def render_chart(chart_obj):
    if chart_obj["type"] == "bar":
        df = chart_obj["data"]
        bar_chart = alt.Chart(df).mark_bar().encode(
            x=alt.X("material:N", sort="-y"),
            y="actual_quantity:Q",
            tooltip=["material", "actual_quantity"]
        ).properties(width=700, height=400)
        st.altair_chart(bar_chart)
