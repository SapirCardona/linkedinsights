import streamlit as st
import pandas as pd
import plotly.express as px

# Title of the app
st.title("LinkedIn Dashboard")

# File uploader to upload Excel file (added unique key)
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"], key="file_uploader_1")

if uploaded_file is not None:
    # Read the Excel file using pandas
    data = pd.read_excel(uploaded_file, sheet_name="FOLLOWERS", header=2)

    # ------------------- Info Cards -------------------
    if uploaded_file is not None:
        # Load the relevant sheets
        discovery_data = pd.read_excel(uploaded_file, sheet_name="DISCOVERY")

        # Extract the required values from the DISCOVERY sheet
        total_impressions = \
        discovery_data.loc[discovery_data['Overall Performance'] == 'Impressions', '4/1/2024 - 3/31/2025'].values[0]
        unique_users = \
        discovery_data.loc[discovery_data['Overall Performance'] == 'Members reached', '4/1/2024 - 3/31/2025'].values[0]

        # Display the values in Streamlit as three beautiful cards
        col1, col2, = st.columns(2)


        with col1:
            st.subheader("Unique Users Exposed")
            st.metric(label="Unique Users", value=unique_users)

        with col2:
            st.subheader("Total Impressions")
            st.metric(label="Impressions", value=total_impressions)

    # Convert 'Date' column to datetime format
    data['Date'] = pd.to_datetime(data['Date'], errors='coerce')  # Convert to datetime with error handling

    # Create a new column for Year-Month
    data['Year-Month'] = data['Date'].dt.to_period('M')  # Extract Year-Month from date

    # Read the "Total followers" value from cell A2
    total_followers = pd.read_excel(uploaded_file, sheet_name="FOLLOWERS", header=None).iloc[0, 0]


    # Group the data by Year-Month and aggregate the sum for each month
    monthly_data = data.groupby('Year-Month').agg({
        'New followers': 'sum',
    }).reset_index()

    # Calculate the growth rate (percentage change) month over month
    monthly_data['Growth Rate (%)'] = monthly_data['New followers'].pct_change() * 100
    # Display the monthly data
    #st.write("Monthly Data", monthly_data)

    # ------------------- Graph 1: New Connections over Time (Bar Chart) -------------------
    monthly_data['Year-Month'] = monthly_data['Year-Month'].astype(str)

    fig1 = px.bar(monthly_data,
                  x='Year-Month',
                  y='New followers',
                  title="New Connections Over Time")

    st.plotly_chart(fig1)

    # ------------------- Graph 2: Monthly Growth Rate of Connections (Line Chart) -------------------
    # Calculate the total number of connections per month (cumulative)
    connections_per_month = data.groupby(data['Date'].dt.to_period('M')).size().cumsum()

    # Calculate the growth rate as: (current month - previous month) * 100
    fig2 = px.line(monthly_data,
                   x='Year-Month',
                   y='Growth Rate (%)',
                   title="Monthly Growth Rate of Followers (%)")
    fig2.update_layout(
        xaxis_title="Month",
        yaxis_title="Growth Rate (%)"
    )

    # Calculate the monthly growth rate (difference between current month and previous month)
    monthly_data['Growth Rate (%)'] = monthly_data['New followers'].diff().fillna(0) / monthly_data[
        'New followers'].shift(1).fillna(1) * 100

    # Adding circles at each data point
    fig2.update_traces(mode='lines+markers',  # This ensures we have both lines and markers (circles)
                       marker=dict(symbol='circle', size=8))  # Add circles at each point

    st.plotly_chart(fig2)

    # ------------------- Graph 3: Most Active Days (Horizontal Bar Chart) -------------------
    if uploaded_file is not None:
        # Read the relevant sheet (TOP POSTS)
        posts_data = pd.read_excel(uploaded_file, sheet_name="TOP POSTS",
                                   header=2)  # Adjust header to start from the third row

        # Extract columns B and F (adjusting for correct rows)
        column_b = posts_data.iloc[2:, 1]  # Start from B3
        column_f = posts_data.iloc[3:, 5]  # Start from F4

        # Merge the two columns (B and F) into one series
        all_dates = pd.concat([column_b, column_f], ignore_index=True)

        # Convert the dates to datetime format
        all_dates = pd.to_datetime(all_dates, errors='coerce')

        # Calculate the day of the week for each date (0=Monday, 6=Sunday)
        all_dates_day = all_dates.dt.day_name()

        # Count the occurrences of each day of the week
        posts_per_day = all_dates_day.value_counts()

        # Reindex to ensure every day of the week is included in the correct order (Sunday first)
        days_of_week = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
        posts_per_day = posts_per_day.reindex(days_of_week, fill_value=0)

        # Display the table with the number of occurrences for each day of the week
        #st.write("Number of Posts Published per Day of the Week:")
        #st.write(posts_per_day)

        # Create a bar chart showing the posts per day
        fig = px.bar(posts_per_day,
                     x=posts_per_day.index,
                     y=posts_per_day,
                     title="Posts Published by Day of the Week",
                     color=posts_per_day)  # Color the bars for better visibility

        # Display the chart in Streamlit
        st.plotly_chart(fig)

    # ------------------- Graph 4: Total Engagements by Month (Bar Chart) -------------------
    # Read data from the "ENGAGEMENT" sheet
    engagement_data = pd.read_excel(uploaded_file, sheet_name="ENGAGEMENT")

    # Convert the 'Date' column to datetime format
    engagement_data['Date'] = pd.to_datetime(engagement_data['Date'], errors='coerce')

    # Add a new column for the day of the week
    engagement_data['Day of Week'] = engagement_data['Date'].dt.day_name()

    # Display the updated DataFrame
    #st.write("Engagement Data with Day of the Week:", engagement_data)

    # Group by 'Day of Week' and sum the 'Impressions' and 'Engagements'
    weekly_engagement_summary = engagement_data.groupby('Day of Week').agg({
        'Impressions': 'sum',
        'Engagements': 'sum'
    }).reindex(['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'])  # Ensure correct order

    # Display the grouped table
    #st.write("Weekly Engagement Summary:", weekly_engagement_summary)

    # Bar chart for Impressions by Day of the Week
    fig_impressions = px.bar(weekly_engagement_summary,
                             x=weekly_engagement_summary.index,
                             y='Impressions',
                             title="Impressions by Day of the Week",
                             labels={'Impressions': 'Count', 'index': 'Day of the Week'},
                             text_auto=True,
                             color_discrete_sequence=['#1f77b4'])  # Blue color for impressions

    # Display the Impressions chart in Streamlit
    st.plotly_chart(fig_impressions)

    # Bar chart for Engagements by Day of the Week
    fig_engagements = px.bar(weekly_engagement_summary,
                             x=weekly_engagement_summary.index,
                             y='Engagements',
                             title="Engagements by Day of the Week",
                             labels={'Engagements': 'Count', 'index': 'Day of the Week'},
                             text_auto=True,
                             color_discrete_sequence=['#ff7f0e'])  # Orange color for engagements

    # Display the Engagements chart in Streamlit
    st.plotly_chart(fig_engagements)

    # ------------------- Top 5 Job Titles -------------------
    if uploaded_file is not None:
        # Load the "DEMOGRAPHICS" sheet from the uploaded file
        data = pd.read_excel(uploaded_file, sheet_name="DEMOGRAPHICS")

        # Filter the data to include only job titles
        job_titles_data = data[data['Top Demographics'] == 'Job titles']

        # Select only relevant columns: job title and percentage
        job_titles_data = job_titles_data[['Value', 'Percentage']]

        # Convert 'Percentage' to string, remove the '%' sign (אם יש), ואז להמיר למספר ולהכפיל ב-100
        job_titles_data['Percentage'] = job_titles_data['Percentage'].astype(str).str.replace('%', '').astype(
            float) * 100

        # Sort job titles by percentage in descending order and select the top 5
        top_job_titles = job_titles_data.sort_values(by='Percentage', ascending=False).head(5)

        # Convert back to percentage format
        top_job_titles['Percentage'] = top_job_titles['Percentage'].apply(lambda x: f"{x:.2f}%")

        # Display the top 5 job titles in Streamlit
        st.subheader("Top 5 Job Titles")
        st.write(top_job_titles)

    # ------------------- Top 5 Seniority Levels -------------------
    # Filter the data to include only seniority
    seniority_data = data[data['Top Demographics'] == 'Seniority']

    # Select only relevant columns: seniority level and percentage
    seniority_data = seniority_data[['Value', 'Percentage']]

    # Convert 'Percentage' to string, remove the '%' sign (אם יש), ואז להמיר למספר ולהכפיל ב-100
    seniority_data['Percentage'] = seniority_data['Percentage'].astype(str).str.replace('%', '').astype(float) * 100

    # Sort seniority by percentage in descending order
    top_seniority = seniority_data.sort_values(by='Percentage', ascending=False)

    # Convert back to percentage format
    top_seniority['Percentage'] = top_seniority['Percentage'].apply(lambda x: f"{x:.2f}%")

    # Display the sorted seniority data in Streamlit
    st.subheader("Top Seniority Levels")
    st.write(top_seniority)


