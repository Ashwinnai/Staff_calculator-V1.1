import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from pyworkforce.queuing import MultiErlangC
import io
import re
import time
import numpy as np

# Define the function to calculate staffing using MultiErlangC
@st.cache_data
def calculate_staffing(awt, shrinkage, max_occupancy, avg_aht, target, calls, interval):
    param_grid = {"transactions": [calls], "aht": [avg_aht / 60], "interval": [30], "asa": [awt / 60], "shrinkage": [shrinkage / 100]}
    multi_erlang = MultiErlangC(param_grid=param_grid, n_jobs=-1)
    required_positions_scenarios = {"service_level": [target / 100], "max_occupancy": [max_occupancy / 100]}
    return multi_erlang.required_positions(required_positions_scenarios)

# Function to sanitize sheet names
def sanitize_sheet_name(sheet_name):
    return re.sub(r'[\\/*?:"<>|]', "_", sheet_name)

# Function to generate Excel report
def generate_excel(all_scenarios, comparison_df, summary_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for scenario_name, (staffing_df, total_staffing) in all_scenarios.items():
            short_name = f'Scn_{scenario_name.split()[1]}'
            short_name_staffing = sanitize_sheet_name(f'{short_name}_Staff')
            short_name_summary = sanitize_sheet_name(f'{short_name}_Sum')

            short_name_staffing = short_name_staffing[:31]
            short_name_summary = short_name_summary[:31]

            staffing_pivot = staffing_df.pivot(index='Interval', columns='Day', values='positions')

            staffing_pivot.to_excel(writer, sheet_name=short_name_staffing, index=True)
            total_staffing.to_excel(writer, sheet_name=short_name_summary)

        # Export Comparison of Scenarios
        comparison_df.to_excel(writer, sheet_name='Comparison of Scenarios', index=True)

        # Export Scenario Summary Table
        summary_df.to_excel(writer, sheet_name='Scenario Summary Table', index=True)

    return output

# Function to adjust interval distribution to ensure it sums to 100% for each day
def adjust_interval_distribution(percentage_df):
    for day in percentage_df.columns:
        total_percentage = percentage_df[day].sum()

        # Calculate the adjusted percentage based on the proportion of each interval's percentage
        percentage_df[day] = percentage_df[day] / total_percentage * 100

        # Ensure the sum of Adjusted Percentage is exactly 100%
        adjusted_total = percentage_df[day].sum()
        difference = 100 - adjusted_total

        # If there's a small difference, adjust the largest interval
        if abs(difference) > 0.001:
            largest_index = percentage_df[day].idxmax()
            percentage_df.at[largest_index, day] += difference

    return percentage_df

# Function to adjust weekly distribution
def adjust_weekly_distribution(distribution_data):
    total_percentage = distribution_data["Percentage"].sum()

    # Calculate the adjusted percentage based on the proportion of each day's percentage
    distribution_data["Adjusted Percentage"] = distribution_data["Percentage"] / total_percentage * 100

    # Ensure the sum of Adjusted Percentage is exactly 100%
    adjusted_total = distribution_data["Adjusted Percentage"].sum()
    difference = 100 - adjusted_total

    # If there's a small difference, adjust the day with the largest percentage
    if abs(difference) > 0.001:
        largest_index = distribution_data["Adjusted Percentage"].idxmax()
        distribution_data.at[largest_index, "Adjusted Percentage"] += difference

    return distribution_data

# Function to validate numeric inputs and convert them to float safely
def validate_and_convert_to_float(value, input_name):
    try:
        return float(value)
    except ValueError:
        raise ValueError(f"Invalid input detected in {input_name}: '{value}' cannot be converted to a number. Please correct the input.")

# Function to validate the inputs before performing calculations
def validate_inputs(*args):
    for i, value in enumerate(args):
        if value is None or isinstance(value, str) or np.isnan(value):
            return False, f"Invalid input detected: Input #{i+1} is not a valid number. Please check your input values."
    return True, None

# Main function of the Streamlit app
def main():
    # Initialize days and intervals
    days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    intervals = pd.date_range("00:00", "23:30", freq="30min").time

    st.set_page_config(page_title="Staffing Calculator", layout="wide")

    # App title
    st.title("Staffing Calculator Multiple Scenario Tester")
    st.sidebar.markdown("### Made by Ashwin Nair")

    # User Guide section in sidebar
    with st.sidebar.expander("User Guide", expanded=False):
        st.markdown("""
        ## Condensed User Guide

        1. **Inputs:** Enter the acceptable waiting time, shrinkage, max occupancy, and service level targets in the sidebar.
        2. **AHT:** Choose to input AHT values either for all intervals and days or at an interval level for each day.
        3. **Volume Distribution:** Enter weekly call volume and adjust the percentage distribution across days and intervals.
        4. **Scenarios:** Click "Calculate Staffing Requirements" to generate staffing scenarios.
        5. **Review:** Review the generated scenarios, staffing requirements, heatmaps, and bar plots.
        6. **Download Report:** Export the results as an Excel report by clicking the "Download Excel" button.

        **Note:** The app retains your scenarios, comparisons, and summaries even after downloading the report.
        """)

    # User input parameters
    with st.sidebar.expander("User Inputs", expanded=False):
        acceptable_waiting_times_input = st.text_input("Acceptable Waiting Time (seconds, comma-separated)", "30")
        try:
            acceptable_waiting_times = [validate_and_convert_to_float(awt, "Acceptable Waiting Times") for awt in acceptable_waiting_times_input.split(',')]
        except ValueError as e:
            st.error(str(e))
            st.stop()
        
        shrinkages_input = st.text_input("Shrinkage (% , comma-separated)", "20")
        try:
            shrinkages = [validate_and_convert_to_float(shrink, "Shrinkage") for shrink in shrinkages_input.split(',')]
        except ValueError as e:
            st.error(str(e))
            st.stop()
        
        max_occupancies_input = st.text_input("Max Occupancy (% , comma-separated)", "80")
        try:
            max_occupancies = [validate_and_convert_to_float(occ, "Max Occupancy") for occ in max_occupancies_input.split(',')]
        except ValueError as e:
            st.error(str(e))
            st.stop()

        service_level_targets_input = st.text_input("Service Level Targets (% , comma-separated)", "80")
        try:
            service_level_targets = [validate_and_convert_to_float(target, "Service Level Targets") for target in service_level_targets_input.split(',')]
        except ValueError as e:
            st.error(str(e))
            st.stop()

        working_hours = st.number_input("Working Hours per Day", min_value=1.0, max_value=24.0, value=8.0, step=0.5)
        working_days = st.number_input("Working Days per Week", min_value=1.0, max_value=7.0, value=5.0, step=0.5)

        aht_input_option = st.radio(
            "Choose AHT (Average Handling Time) Input Method:",
            ("Multiple AHT values for all intervals and days", "AHT table at interval level for each day")
        )

        if aht_input_option == "Multiple AHT values for all intervals and days":
            average_handling_times_input = st.text_input("Average Handling Times (seconds, comma-separated)", "300")
            try:
                average_handling_times = [validate_and_convert_to_float(aht, "Average Handling Times") for aht in average_handling_times_input.split(',')]
            except ValueError as e:
                st.error(str(e))
                st.stop()
        else:
            average_handling_times = []

    manual_input = st.sidebar.checkbox("Auto Input Calls Offered per Interval", value=False)

    if manual_input:
        with st.sidebar.expander("Weekly Volume", expanded=True):
            weekly_volume = st.number_input("Enter Weekly Volume", min_value=0, value=800, step=100)

        st.header("Volume Weekly Distribution (%) & Percentage Distribution by Interval")
        with st.expander("Weekly Distribution (%)", expanded=True):
            st.markdown("### Input Daily Distribution (%) for Each Day")
            distribution_data = pd.DataFrame({"Day": days, "Percentage": [15.0, 20.0, 20.0, 15.0, 15.0, 10.0, 5.0]})
            distribution_data["Percentage"] = distribution_data["Percentage"].astype(float)
            distribution_data = st.data_editor(distribution_data, key="distribution_data_editor")

            adjusted_distribution_data = adjust_weekly_distribution(distribution_data)
            st.markdown(f"**Sum of Adjusted Weekly Distribution: {adjusted_distribution_data['Adjusted Percentage'].sum()}%**")
            st.dataframe(adjusted_distribution_data)

        with st.expander("Percentage Distribution by Interval (Sunday to Saturday)", expanded=True):
            st.markdown("### Input Percentage Distribution for Each 30-minute Interval")
            initial_data = [[2.08 for _ in days] for _ in intervals]
            percentage_df = pd.DataFrame(initial_data, index=intervals, columns=days)
            percentage_df = st.data_editor(percentage_df, key="percentage_df_editor")

            adjusted_percentage_df = adjust_interval_distribution(percentage_df)

            sum_columns = adjusted_percentage_df.sum()
            sum_df = pd.DataFrame(sum_columns).T
            sum_df.index = ['Sum']

            st.markdown("### Sum of Interval Distribution for Each Day")
            st.dataframe(sum_df)

            st.markdown("### Interval Distribution Data")
            st.dataframe(adjusted_percentage_df)

            st.markdown("### Sum of All Intervals for Each Day")
            interval_sums_df = adjusted_percentage_df.sum(axis=0).to_frame(name='Total for Each Day')
            st.dataframe(interval_sums_df)

        st.header("Calls Offered per 30-minute Interval (Sunday to Saturday)")
        calls_df = pd.DataFrame(index=intervals)
        daily_volumes = [(weekly_volume * (dist / 100)) for dist in adjusted_distribution_data["Adjusted Percentage"].tolist()]

        for day, daily_volume in zip(days, daily_volumes):
            total_percentage = sum(adjusted_percentage_df[day])
            try:
                calls_per_interval = [validate_and_convert_to_float(daily_volume * (p / 100), f"Calls for {day} during {interval}") for p, interval in zip(adjusted_percentage_df[day], intervals)]
            except ValueError as e:
                st.error(str(e))
                st.stop()
            calls_df[day] = calls_per_interval

        st.session_state["calls_df"] = st.data_editor(calls_df, key="manual_calls_df_editor")

    else:
        st.header("Intraday Volume DataFrame")
        if "calls_df" not in st.session_state:
            st.session_state["calls_df"] = pd.DataFrame(index=intervals, columns=days)
        st.session_state["calls_df"] = st.data_editor(st.session_state["calls_df"], key="calls_df_editor")

    if aht_input_option == "AHT table at interval level for each day":
        st.header("Average Handling Time (AHT) per 30-minute Interval (Sunday to Saturday)")
        data_aht = {day: [0.0] * len(intervals) for day in days}
        aht_df = pd.DataFrame(data_aht, index=intervals)
        
        if "aht_df" not in st.session_state:
            st.session_state["aht_df"] = aht_df
        st.session_state["aht_df"] = st.data_editor(st.session_state["aht_df"], key="aht_df_editor")
        aht_df = st.session_state["aht_df"]

    if st.button("Calculate Staffing Requirements"):
        start_time = time.time()  # Start timing
        progress_bar = st.progress(0)
        total_combinations = (
            len(acceptable_waiting_times)
            * len(shrinkages)
            * len(max_occupancies)
            * len(service_level_targets)
            * len(days)
            * len(intervals)
        )
        current_progress = 0

        scenario_number = 1
        all_scenarios = st.session_state.get("all_scenarios", {})  # Retain previous scenarios

        for awt in acceptable_waiting_times:
            for shrinkage in shrinkages:
                for max_occupancy in max_occupancies:
                    if aht_input_option == "Multiple AHT values for all intervals and days":
                        for avg_aht in average_handling_times:
                            for target in service_level_targets:
                                scenario_name = f"Scenario {scenario_number}: AWT={awt}s, Shrinkage={shrinkage}%, Max Occupancy={max_occupancy}%, AHT={avg_aht}s, SLT={target}%"
                                with st.expander(scenario_name, expanded=False):
                                    staffing_results = []
                                    error_displayed = False  # Flag to track if an error was already shown
                                    for day in days:
                                        for interval, calls in zip(intervals, st.session_state["calls_df"][day]):
                                            if calls == 0:
                                                continue

                                            try:
                                                calls = validate_and_convert_to_float(calls, f"Calls for {day} during {interval}")
                                            except ValueError as e:
                                                if not error_displayed:  # Display the error only once
                                                    st.error(str(e))
                                                    error_displayed = True
                                                break  # Exit the loop to prevent further processing

                                            # Validate inputs before calculating staffing
                                            is_valid, error_message = validate_inputs(awt, shrinkage, max_occupancy, avg_aht, target, calls)
                                            if not is_valid:
                                                if not error_displayed:  # Display the error only once
                                                    st.error(error_message)
                                                    error_displayed = True
                                                break  # Exit the loop to prevent further processing

                                            positions_requirements = calculate_staffing(awt, shrinkage, max_occupancy, avg_aht, target, calls, interval)

                                            for requirement in positions_requirements:
                                                requirement.update({
                                                    "Day": day,
                                                    "Interval": interval,
                                                    "AWT": awt,
                                                    "Shrinkage": shrinkage,
                                                    "Max Occupancy": max_occupancy,
                                                    "Average AHT": avg_aht,
                                                    "Service Level Target": target,
                                                    "raw_positions": requirement.get("raw_positions", 0),
                                                    "positions": requirement.get("positions", 0),
                                                    "service_level": requirement.get("service_level", 0),
                                                    "occupancy": requirement.get("occupancy", 0),
                                                    "waiting_probability": requirement.get("waiting_probability", 0)
                                                })
                                                staffing_results.append(requirement)

                                            current_progress += 1
                                            progress_bar.progress(min(current_progress / total_combinations, 1.0))

                                    staffing_df = pd.DataFrame(staffing_results)
                                    
                                    required_columns = [
                                        "Day", "Interval", "AWT", "Shrinkage", "Max Occupancy",
                                        "Average AHT", "Service Level Target", "raw_positions", "positions",
                                        "service_level", "occupancy", "waiting_probability"
                                    ]

                                    for column in required_columns:
                                        if column not in staffing_df.columns:
                                            staffing_df[column] = 0

                                    staffing_df = staffing_df[required_columns]

                                    st.write(f"Staffing Requirements for AWT: {awt}s, Shrinkage: {shrinkage}%, Max Occupancy: {max_occupancy}%, Average AHT: {avg_aht}s, Service Level Target: {target}%")
                                    st.dataframe(staffing_df)

                                    total_staffing = staffing_df.groupby("Day")[["raw_positions", "positions"]].sum()
                                    total_staffing["Sum of Raw Positions"] = total_staffing["raw_positions"]
                                    total_staffing["Sum of Positions"] = total_staffing["positions"]
                                    total_staffing["Required staff without Peak staffing"] = total_staffing["Sum of Positions"] / 2 / working_hours
                                    total_staffing["Peak Staffing requirement"] = total_staffing["Required staff without Peak staffing"].max()
                                    total_staffing["Sum of the Week"] = total_staffing["Required staff without Peak staffing"].sum()
                                    total_staffing["Required staff without Peak staffing"] = total_staffing["Sum of the Week"] / working_days

                                    total_staffing = total_staffing.reindex(["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"])
                                    
                                    st.write("Required Staffing")
                                    st.dataframe(total_staffing)

                                    st.write("Interactive Heatmap of Required Staffing Levels")
                                    heatmap_data = staffing_df.pivot_table(index="Day", columns="Interval", values="positions", aggfunc="mean")
                                    heatmap_data = heatmap_data.reindex(["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"])
                                    fig = go.Figure(data=go.Heatmap(
                                        z=heatmap_data.values,
                                        x=heatmap_data.columns,
                                        y=heatmap_data.index,
                                        colorscale='YlGnBu'
                                    ))
                                    fig.update_layout(
                                        title=f'Staffing Levels Heatmap (AWT: {awt}s, Shrinkage={shrinkage}%, Max Occupancy={max_occupancy}%, Average AHT={avg_aht}s, Service Level Target={target}%)',
                                        xaxis_nticks=48
                                    )
                                    st.plotly_chart(fig)

                                    st.write("Interactive Bar Plot of Staffing Levels for Each Day")
                                    for day in days:
                                        bar_plot = go.Figure()
                                        bar_plot.add_trace(go.Bar(
                                            x=intervals,
                                            y=staffing_df[staffing_df["Day"] == day]["positions"],
                                            name=day
                                        ))
                                        bar_plot.update_layout(
                                            title=f'Staffing Levels Bar Plot for {day} (AWT: {awt}s, Shrinkage: {shrinkage}%, Max Occupancy={max_occupancy}%, AHT={avg_aht}s, SLT={target}%)',
                                            xaxis_title="Interval",
                                            yaxis_title="Staffing Level",
                                            xaxis_nticks=48
                                        )
                                        st.plotly_chart(bar_plot)

                                    all_scenarios[scenario_name] = (staffing_df, total_staffing)
                                    
                                scenario_number += 1

                    else:
                        for target in service_level_targets:
                            scenario_name = f"Scenario {scenario_number}: AWT={awt}s, Shrinkage={shrinkage}%, Max Occupancy={max_occupancy}%, SLT={target}%"
                            with st.expander(scenario_name, expanded=False):
                                staffing_results = []
                                error_displayed = False  # Flag to track if an error was already shown
                                for day in days:
                                    for interval, calls, aht in zip(intervals, st.session_state["calls_df"][day], aht_df[day]):
                                        if calls == 0 or aht == 0:
                                            continue

                                        try:
                                            calls = validate_and_convert_to_float(calls, f"Calls for {day} during {interval}")
                                            aht = validate_and_convert_to_float(aht, f"AHT for {day} during {interval}")
                                        except ValueError as e:
                                            if not error_displayed:  # Display the error only once
                                                st.error(str(e))
                                                error_displayed = True
                                            break  # Exit the loop to prevent further processing

                                        # Validate inputs before calculating staffing
                                        is_valid, error_message = validate_inputs(awt, shrinkage, max_occupancy, aht, target, calls)
                                        if not is_valid:
                                            if not error_displayed:  # Display the error only once
                                                st.error(error_message)
                                                error_displayed = True
                                            break  # Exit the loop to prevent further processing

                                        positions_requirements = calculate_staffing(awt, shrinkage, max_occupancy, aht, target, calls, interval)

                                        for requirement in positions_requirements:
                                            requirement.update({
                                                "Day": day,
                                                "Interval": interval,
                                                "AWT": awt,
                                                "Shrinkage": shrinkage,
                                                "Max Occupancy": max_occupancy,
                                                "Average AHT": aht,
                                                "Service Level Target": target,
                                                "raw_positions": requirement.get("raw_positions", 0),
                                                "positions": requirement.get("positions", 0),
                                                "service_level": requirement.get("service_level", 0),
                                                "occupancy": requirement.get("occupancy", 0),
                                                "waiting_probability": requirement.get("waiting_probability", 0)
                                            })
                                            staffing_results.append(requirement)

                                        current_progress += 1
                                        progress_bar.progress(min(current_progress / total_combinations, 1.0))

                                staffing_df = pd.DataFrame(staffing_results)

                                required_columns = [
                                    "Day", "Interval", "AWT", "Shrinkage", "Max Occupancy",
                                    "Average AHT", "Service Level Target", "raw_positions", "positions",
                                    "service_level", "occupancy", "waiting_probability"
                                ]

                                for column in required_columns:
                                    if column not in staffing_df.columns:
                                        staffing_df[column] = 0

                                staffing_df = staffing_df[required_columns]

                                st.write(f"Staffing Requirements for AWT: {awt}s, Shrinkage: {shrinkage}%, Max Occupancy: {max_occupancy}%, Average AHT: {aht}s, Service Level Target: {target}%")
                                st.dataframe(staffing_df)

                                total_staffing = staffing_df.groupby("Day")[["raw_positions", "positions"]].sum()
                                total_staffing["Sum of Raw Positions"] = total_staffing["raw_positions"]
                                total_staffing["Sum of Positions"] = total_staffing["positions"]
                                total_staffing["Required staff without Peak staffing"] = total_staffing["Sum of Positions"] / 2 / working_hours
                                total_staffing["Peak Staffing requirement"] = total_staffing["Required staff without Peak staffing"].max()
                                total_staffing["Sum of the Week"] = total_staffing["Required staff without Peak staffing"].sum()
                                total_staffing["Required staff without Peak staffing"] = total_staffing["Sum of the Week"] / working_days

                                total_staffing = total_staffing.reindex(["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"])
                                
                                st.write("Required Staffing")
                                st.dataframe(total_staffing)

                                st.write("Interactive Heatmap of Required Staffing Levels")
                                heatmap_data = staffing_df.pivot_table(index="Day", columns="Interval", values="positions", aggfunc="mean")
                                heatmap_data = heatmap_data.reindex(["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"])
                                fig = go.Figure(data=go.Heatmap(
                                    z=heatmap_data.values,
                                    x=heatmap_data.columns,
                                    y=heatmap_data.index,
                                    colorscale='YlGnBu'
                                ))
                                fig.update_layout(
                                    title=f'Staffing Levels Heatmap (AWT: {awt}s, Shrinkage={shrinkage}%, Max Occupancy={max_occupancy}%, Average AHT={aht}s, Service Level Target={target}%)',
                                        xaxis_nticks=48
                                    )
                                st.plotly_chart(fig)

                                st.write("Interactive Bar Plot of Staffing Levels for Each Day")
                                for day in days:
                                    bar_plot = go.Figure()
                                    bar_plot.add_trace(go.Bar(
                                        x=intervals,
                                        y=staffing_df[staffing_df["Day"] == day]["positions"],
                                        name=day
                                    ))
                                    bar_plot.update_layout(
                                        title=f'Staffing Levels Bar Plot for {day} (AWT: {awt}s, Shrinkage: {shrinkage}%, Max Occupancy={max_occupancy}%, AHT={aht}s, SLT={target}%)',
                                        xaxis_title="Interval",
                                        yaxis_title="Staffing Level",
                                        xaxis_nticks=48
                                    )
                                    st.plotly_chart(bar_plot)

                                all_scenarios[scenario_name] = (staffing_df, total_staffing)

                            scenario_number += 1

        st.session_state["all_scenarios"] = all_scenarios  # Store scenarios in session state

        progress_bar.empty()
# End timing
        end_time = time.time()  
        total_time = end_time - start_time  
          # Calculate total time taken
        st.write(f"**Total Time Taken for All Scenarios:** {total_time:.2f} seconds")

        if all_scenarios:
            st.header("Comparison of Scenarios")
            comparison_df = pd.DataFrame()

            for scenario, (staffing_df, total_staffing) in all_scenarios.items():
                summary = total_staffing[["Required staff without Peak staffing", "Peak Staffing requirement"]].rename(
                    columns={"Required staff without Peak staffing": "Required staff without Peak staffing", "Peak Staffing requirement": "Peak Staffing requirement"}
                )
                summary["Scenario"] = scenario
                comparison_df = pd.concat([comparison_df, summary])

            st.dataframe(comparison_df)

            st.header("Scenario Summary Table")
            summary_df = pd.DataFrame()

            for scenario, (staffing_df, total_staffing) in all_scenarios.items():
                required_staff_without_peak = total_staffing["Required staff without Peak staffing"].max()  # Use max instead of sum
                peak_staffing_requirement = total_staffing["Peak Staffing requirement"].max()
                
                temp_df = pd.DataFrame({
                    "Scenario": [scenario],
                    "Required staff without Peak staffing": [required_staff_without_peak],
                    "Peak Staffing requirement": [peak_staffing_requirement]
                })
                
                summary_df = pd.concat([summary_df, temp_df], ignore_index=True)

            st.dataframe(summary_df)

            # Export report
            st.markdown("## Download Your Report")
            excel_data = generate_excel(all_scenarios, comparison_df, summary_df)
            st.download_button(
                label="Download Excel",
                data=excel_data.getvalue(),
                file_name='staffing_requirements.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

if __name__ == "__main__":
    main()
