import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
from itertools import zip_longest

# Set page configuration
st.set_page_config(layout="wide")
st.title("📊 CU Analysis App")

# File uploader
uploaded_file = st.file_uploader("📂 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Read Excel file
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        sheets_data = {}

        # Process each sheet
        for sheet in sheet_names:
            try:
                df = pd.read_excel(uploaded_file, sheet_name=sheet)

                # Drop first column
                df = df.drop(df.columns[0], axis=1)

                # Drop first row (junk header), reset index
                df = df.drop(index=0).reset_index(drop=True)

                # Set second row as header
                df.columns = df.iloc[0]
                df = df.drop(index=0).reset_index(drop=True)

                # Drop rows with all NaN
                df = df.dropna(how='all')

                # Add client name
                df['client'] = sheet

                # Save to dictionary
                sheets_data[sheet] = df

                st.success(f"✅ Processed: {sheet} | Rows kept: {len(df)}")

            except Exception as e:
                st.error(f"❌ Error processing sheet '{sheet}': {e}")

        # Clean column names
        for sheet in sheets_data:
            sheets_data[sheet].columns = (
                sheets_data[sheet].columns
                .str.strip()  # Remove leading/trailing spaces
                .str.lower()  # Convert to lowercase
                .str.replace(r'\s+', ' ', regex=True)  # Normalize inner spaces
            )

        # Combine all sheets
        combined_df = pd.concat(sheets_data.values(), ignore_index=True)
        combined_df = combined_df.dropna(how='all')

        # Forward fill 'job no'
        combined_df['job no'] = combined_df['job no'].ffill()

        # Convert 'quantity/mt' and 'invoice amount' to numeric
        for col in ['quantity/mt', 'invoice amount']:
            if col in combined_df.columns:
                combined_df[col] = (
                    combined_df[col]
                    .astype(str)
                    .str.replace(',', '', regex=False)
                    .str.replace('\xa0', '', regex=False)
                    .str.strip()
                    .replace('', '0')
                    .astype(float)
                )
            else:
                st.warning(f"⚠️ Column '{col}' not found in combined_df")

        # Summary by client and currency
        summary = combined_df.groupby(['client', 'currency']).agg({
            'quantity/mt': 'sum',
            'invoice amount': 'sum'
        }).reset_index()

        # Drop rows where currency is blank AND both values are 0
        summary = summary[
            ~((summary['currency'].astype(str).str.strip() == '') &
              (summary['quantity/mt'] == 0) &
              (summary['invoice amount'] == 0))
        ]

        summary = summary.dropna(subset=['quantity/mt', 'invoice amount'], how='all')

        # Display summary
        st.subheader("Summary by Client and Currency")
        st.dataframe(
            summary.style.format({
                'quantity/mt': '{:,.2f}',
                'invoice amount': '{:,.2f}'
            }),
            use_container_width=True
        )

        # Job No by Month Table
        st.subheader("Clients by Job No")
        first_jobs = (
            combined_df[['client', 'job no']]
            .dropna()
            .drop_duplicates('client')
            .copy()
        )

        # Extract month from 'job no'
        first_jobs = first_jobs[first_jobs['job no'].astype(str).str.count('\.') == 2]
        first_jobs['Month'] = first_jobs['job no'].astype(str).str.split('.').str[1]

        # Group clients by Job No month
        month_groups = first_jobs.groupby('Month')['client'].apply(list).to_dict()
        sorted_months = sorted(month_groups)

        # Build equal-height DataFrame
        rows = list(zip_longest(*[month_groups[m] for m in sorted_months], fillvalue=None))
        pivot_df = pd.DataFrame(rows, columns=[f'Job No.{m}' for m in sorted_months])
        pivot_df = pivot_df.dropna(how='all')

        # Display pivot table
        st.dataframe(pivot_df, use_container_width=True)

        # Bar Plots by Client
        st.subheader("Quantity by Commodity (Bar Plots)")
        clients = combined_df['client'].unique()

        for client in clients:
            client_df = combined_df[combined_df['client'] == client]
            summary = (
                client_df
                .groupby('commodity', as_index=False)['quantity/mt']
                .sum()
                .sort_values(by='quantity/mt', ascending=False)
            )

            if summary.empty:
                st.warning(f"⚠️ No data for client: {client}")
                continue

            # Normalize values for color mapping
            norm = plt.Normalize(summary['quantity/mt'].min(), summary['quantity/mt'].max())
            colors = sns.color_palette("viridis", len(summary))

            # Plot
            fig, ax = plt.subplots(figsize=(12, 6))
            bars = ax.bar(summary['commodity'], summary['quantity/mt'], color=colors)

            # Labels
            ax.set_xlabel('Commodity', fontsize=12)
            ax.set_ylabel('Total Quantity (MT)', fontsize=12)
            ax.set_title(f'{client}: Total Quantity by Commodity', fontsize=15, weight='bold')
            ax.tick_params(axis='x', rotation=20, labelsize=10)
            ax.tick_params(axis='y', labelsize=10)

            # Add values on top of bars
            for bar in bars:
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width()/2, height, f'{height:,.0f}', 
                        ha='center', va='bottom', fontsize=9)

            plt.tight_layout()
            ax.grid(axis='y', linestyle='--', alpha=0.5)
            st.pyplot(fig)

        # Pie Charts by Client
        st.subheader("Quantity Distribution by Commodity (Pie Charts)")
        color_palette = px.colors.qualitative.Set3

        for client in clients:
            client_df = combined_df[combined_df['client'] == client]
            summary = (
                client_df
                .groupby('commodity', as_index=False)['quantity/mt']
                .sum()
                .sort_values(by='quantity/mt', ascending=False)
            )

            if summary.empty:
                st.warning(f"⚠️ No data for client: {client}")
                continue

            # Create pie chart
            fig = px.pie(
                summary,
                names='commodity',
                values='quantity/mt',
                title=f'{client} – Quantity Distribution by Commodity',
                color_discrete_sequence=color_palette,
                hover_data={'quantity/mt': True, 'commodity': False}
            )

            fig.update_traces(
                textinfo='label+percent',
                hovertemplate='<b>%{label}</b><br>Quantity: %{value:,.2f} MT<extra></extra>',
                pull=[0.05] * len(summary)
            )

            fig.update_layout(
                height=600,
                width=800,
                title_font_size=20,
                legend_title_text='Commodity',
                legend_font_size=12
            )

            st.plotly_chart(fig, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("Please upload an Excel file to proceed.")

# ✅ Show overall totals if combined_df is ready
if 'combined_df' in locals() and not combined_df.empty:

    st.subheader("📈 Overall Totals Across All Clients")

    # 📦 Overall quantity (all currencies together)
    total_qty = combined_df['quantity/mt'].sum()

    # 💰 Invoice totals per currency
    inv_by_cur = (
        combined_df
        .groupby('currency')['invoice amount']
        .sum()
        .rename(lambda c: c.strip().upper() if isinstance(c, str) else c)
    )

    # 🔢 Metric cards: one for quantity + two for invoice (USD & EGP)
    col1, col2, col3 = st.columns(3)

    col1.metric("📦 Total Quantity (MT)", f"{total_qty:,.2f}")

    col2.metric("💰 Invoice Amount – USD", f"{inv_by_cur.get('USD', 0):,.2f}")

    col3.metric("💰 Invoice Amount – EGP", f"{inv_by_cur.get('EGP', 0):,.2f}")



