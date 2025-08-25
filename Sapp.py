import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import datetime
import re
import math
import itertools
import copy
from ortools.sat.python import cp_model
from pyworkforce.queuing import ErlangC, MultiErlangC
import json
import os
import time
from io import StringIO

# ------------------------------------------------------------------------------
#                           CONFIGURATION & INITIALIZATION
# ------------------------------------------------------------------------------
st.set_page_config(page_title="Advanced WFM Scheduler", layout="wide")
st.title("Multi-Channel WFM: Staffing, Scheduling & Costing")

# --- Initialize Session State ---
# Using .get() with a default value for safe initialization
if "interval_freq" not in st.session_state:
    st.session_state["interval_freq"] = "30min"
if "intervals" not in st.session_state:
    st.session_state["intervals"] = [t.time() for t in pd.date_range("00:00", "23:30", freq="30min")]
if "all_scenarios" not in st.session_state:
    st.session_state['all_scenarios'] = {}
if "scenario_summary" not in st.session_state:
    st.session_state['scenario_summary'] = pd.DataFrame()
if "scheduling_solutions" not in st.session_state:
    st.session_state['scheduling_solutions'] = {}
if 'week_start_day' not in st.session_state:
    st.session_state.week_start_day = "Sunday"
if 'adjusted_headcounts' not in st.session_state:
    st.session_state.adjusted_headcounts = {}

# --- State for Blended Scenarios ---
if 'blended_volumes' not in st.session_state:
    st.session_state.blended_volumes = {}

# --- State for What-If Simulations ---
if 'what_if_simulations' not in st.session_state:
    st.session_state.what_if_simulations = {}
if 'active_what_if_simulation' not in st.session_state:
    st.session_state.active_what_if_simulation = None

# --- Default shifts now have duration, not unpaid break ---
if 'shifts_df' not in st.session_state:
    st.session_state.shifts_df = pd.DataFrame([
        {'Shift Name': 'Morning', 'Start Time': datetime.time(8, 0), 'Shift Length (hours)': 9.0},
        {'Shift Name': 'Mid-Day', 'Start Time': datetime.time(10, 0), 'Shift Length (hours)': 9.0},
        {'Shift Name': 'Evening', 'Start Time': datetime.time(14, 0), 'Shift Length (hours)': 9.0}
    ])

# --- State for Advanced Shift Optimization ---
if 'duration_rules' not in st.session_state:
    st.session_state.duration_rules = {9.0: {'min_days': 5, 'max_days': 5, 'min_off': 2}}
if 'allowed_durations' not in st.session_state:
    st.session_state.allowed_durations = [9.0]
if 'distribution_caps' not in st.session_state:
    st.session_state.distribution_caps = {}
if 'shift_consistency_opt' not in st.session_state:
    st.session_state.shift_consistency_opt = True

# --- NEW: Daily Operational Hours for Shift Optimization ---
if 'daily_op_hours' not in st.session_state:
    st.session_state.daily_op_hours = {
        "Sunday": {"Start Time": datetime.time(0, 0), "End Time": datetime.time(0, 0)}, # Example of 24/7
        "Monday": {"Start Time": datetime.time(0, 0), "End Time": datetime.time(0, 0)}, # Example of 24/7
        "Tuesday": {"Start Time": datetime.time(8, 0), "End Time": datetime.time(20, 0)},
        "Wednesday": {"Start Time": datetime.time(8, 0), "End Time": datetime.time(20, 0)},
        "Thursday": {"Start Time": datetime.time(8, 0), "End Time": datetime.time(20, 0)},
        "Friday": {"Start Time": datetime.time(8, 0), "End Time": datetime.time(22, 0)},
        "Saturday": {"Start Time": datetime.time(9, 0), "End Time": datetime.time(17, 0)},
    }


DAYS_OF_WEEK_OPTIONS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

# ------------------------------------------------------------------------------
#                           HELPER / UTILITY FUNCTIONS
# ------------------------------------------------------------------------------

def safe_update_dataframe(df_key, new_columns, index_ref):
    """
    Updates a DataFrame in session state. If columns change, it preserves
    data from overlapping columns instead of resetting completely.
    """
    old_df = st.session_state.get(df_key)
    new_df = pd.DataFrame(0.0, index=index_ref, columns=new_columns)

    if isinstance(old_df, pd.DataFrame):
        common_cols = old_df.columns.intersection(new_columns)
        if not common_cols.empty:
            # Align indexes to prevent NaN values if indexes differ slightly
            aligned_old_df, _ = old_df.align(new_df, join='right', axis=0)
            new_df[common_cols] = aligned_old_df[common_cols].fillna(0.0)
    
    st.session_state[df_key] = new_df

def format_duration(seconds):
    """Formats a duration in seconds into a human-readable string."""
    if seconds < 60:
        return f"{seconds:.2f} seconds"
    else:
        minutes = int(seconds // 60)
        remaining_seconds = seconds % 60
        return f"{minutes} minute(s) and {remaining_seconds:.2f} seconds"

@st.cache_data
def get_week_start(d, start_day_name="Sunday"):
    day_map = {"Sunday": 0, "Monday": 1, "Tuesday": 2, "Wednesday": 3, "Thursday": 4, "Friday": 5, "Saturday": 6}
    start_day_idx = day_map.get(start_day_name, 0)
    # Convert date 'd' to datetime if it's not already
    if isinstance(d, datetime.date) and not isinstance(d, datetime.datetime):
         d = datetime.datetime.combine(d, datetime.datetime.min.time())
    date_day_of_week_idx = (d.weekday() + 1) % 7
    days_to_subtract = (date_day_of_week_idx - start_day_idx + 7) % 7
    return (d - datetime.timedelta(days=days_to_subtract)).date()

@st.cache_data
def sanitize_name(name):
    return re.sub(r'[\\/*?:"<>|]', "_", name)

def validate_and_convert_to_float(value, input_name):
    if value is None or value == '': return 0.0
    try:
        result = float(value)
        return 0.0 if np.isnan(result) else result
    except (ValueError, TypeError):
        raise ValueError(f"Invalid input '{value}' for {input_name}.")

def calculate_adherence_metrics(req_matrix, sched_matrix, cap_percent, day_order):
    """
    Calculates detailed adherence metrics using floating-point logic to match Excel.
    """
    num_days, num_intervals = req_matrix.shape
    # Ensure intervals_str is a list of formatted strings
    intervals_str = [t.strftime('%H:%M') for t in st.session_state.intervals]


    detail_rows = []

    # 1. Calculate interval-level metrics first
    for d_idx, day_name in enumerate(day_order):
        for p_idx, interval_time in enumerate(intervals_str):
            req = req_matrix[d_idx, p_idx]
            sched = sched_matrix[d_idx, p_idx]

            raw_adherence = 0.0
            if req > 0:
                raw_adherence = (sched / req) * 100

            capped_sched_contribution = 0.0
            if req > 0:
                max_float_contribution = req * cap_percent / 100.0
                capped_sched_contribution = min(float(sched), max_float_contribution)

            capped_adherence = 0.0
            if req > 0:
                capped_adherence = (capped_sched_contribution / req) * 100

            detail_rows.append({
                "Day": day_name,
                "Interval": interval_time,
                "Required": int(req),
                "Scheduled": int(sched),
                "Raw Adherence (%)": raw_adherence,
                "Capped Adherence (%)": capped_adherence,
                "Capped Scheduled Contribution": capped_sched_contribution
            })

    adherence_df = pd.DataFrame(detail_rows)

    # 2. Calculate Daily Adherence based on capped contributions
    daily_summary = adherence_df.groupby('Day').agg(
        Total_Required=('Required', 'sum'),
        Total_Capped_Scheduled_Contribution=('Capped Scheduled Contribution', 'sum')
    ).reindex(day_order)

    daily_adherence_values = {}
    for day_name, row in daily_summary.iterrows():
        total_req = row['Total_Required']
        total_capped_contrib = row['Total_Capped_Scheduled_Contribution']

        if total_req > 0:
            adherence = (total_capped_contrib / total_req) * 100
        else:
            adherence = 0.0

        daily_adherence_values[day_name] = adherence

    # 3. Calculate Weekly Adherence (a weighted average of the daily scores)
    weekly_total_req = daily_summary['Total_Required'].sum()
    weekly_total_capped_contrib = daily_summary['Total_Capped_Scheduled_Contribution'].sum()

    weekly_adherence_value = 0.0
    if weekly_total_req > 0:
        weekly_adherence_value = (weekly_total_capped_contrib / weekly_total_req) * 100
    else:
        weekly_adherence_value = 0.0

    return {
        "weekly_adherence": weekly_adherence_value,
        "daily_adherence": daily_adherence_values,
        "adherence_df": adherence_df
    }

def calculate_fte_metrics_from_matrix(matrix, working_hours, working_days):
    """Calculates average and peak FTE from a requirement matrix."""
    if not isinstance(matrix, list) or not matrix or not working_hours or not working_days:
        return {'avg_fte': 0, 'peak_fte': 0}

    num_intervals_per_day = len(matrix[0]) if matrix else 0
    if num_intervals_per_day == 0:
        return {'avg_fte': 0, 'peak_fte': 0}

    interval_duration_hours = 24 / num_intervals_per_day

    total_required_hours = sum(sum(day) for day in matrix) * interval_duration_hours

    # Avg FTE based on weekly hours
    avg_fte = total_required_hours / (working_hours * working_days) if (working_hours * working_days) > 0 else 0

    # Peak FTE based on the day with the most required hours
    daily_hours = [sum(day) * interval_duration_hours for day in matrix]
    peak_fte_day_hours = max(daily_hours) if daily_hours else 0
    peak_fte = peak_fte_day_hours / working_hours if working_hours > 0 else 0

    return {'avg_fte': avg_fte, 'peak_fte': peak_fte}


def expand_shifts_for_solver(shifts_df):
    """
    Processes user-defined shifts (start time, duration) into
    coverage arrays for the solver. All shift time is considered paid.
    """
    virtual_shifts = []
    shift_groups = {}
    solver_id = 0
    num_intervals = 48

    for _, s_row in shifts_df.iterrows():
        original_name = s_row['Shift Name']
        if not original_name or pd.isna(original_name):
            continue

        start_t = s_row['Start Time']
        duration_h = s_row['Shift Length (hours)']

        if pd.isna(start_t) or pd.isna(duration_h):
            st.warning(f"Shift '{original_name}' is missing a Start Time or Duration and will be skipped.")
            continue

        # Calculate start and end intervals, handling overnight shifts
        start_datetime = datetime.datetime.combine(datetime.date.today(), start_t)
        end_datetime = start_datetime + datetime.timedelta(hours=float(duration_h))

        start_interval = start_t.hour * 2 + start_t.minute // 30
        end_interval = end_datetime.hour * 2 + end_datetime.minute // 30

        # Create base coverage for the entire shift duration
        coverage = [0] * num_intervals
        current_interval = start_interval
        while current_interval != end_interval:
            coverage[current_interval] = 1
            current_interval = (current_interval + 1) % num_intervals

        # For this model, an agent is available and paid for the entire shift duration.
        # Paid breaks are assumed to be part of the overall shrinkage percentage.
        virtual_shifts.append({
            'solver_id': solver_id,
            'display_name': original_name,
            'original_name': original_name,
            'availability_coverage': coverage,
            'payable_coverage': coverage  # Both are the same
        })

        if original_name not in shift_groups:
            shift_groups[original_name] = []
        shift_groups[original_name].append(solver_id)
        solver_id += 1

    return virtual_shifts, shift_groups

@st.cache_data
def process_manual_requirements(req_df, week_start_day_name, working_days_per_week, working_hours_per_day, day_order):
    """
    Processes a DataFrame of manual requirements into a list of weekly jobs for the solver.
    """
    if req_df.empty:
        return []

    req_df = req_df.apply(pd.to_numeric, errors='coerce').fillna(0)

    num_intervals = len(req_df)
    interval_duration_hours = 24 / num_intervals

    weeks = {}
    for date_str in req_df.columns:
        date_obj = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
        week_start = get_week_start(date_obj, week_start_day_name)
        if week_start not in weeks:
            weeks[week_start] = {'data': {}, 'total_req_hours': 0}

        req_values = req_df[date_str].fillna(0).astype(int).tolist()
        weeks[week_start]['data'][date_obj] = req_values
        weeks[week_start]['total_req_hours'] += sum(req_values) * interval_duration_hours

    results = []
    for week_start_date, week_data in weeks.items():
        avg_fte = 0
        if working_days_per_week > 0 and working_hours_per_day > 0:
            avg_fte = week_data['total_req_hours'] / (working_days_per_week * working_hours_per_day)

        week_matrix = []
        # Use the provided day_order to construct the matrix correctly
        for i, day_name in enumerate(day_order):
            current_day = week_start_date + datetime.timedelta(days=i)
            # Find the date_obj in week_data that corresponds to this day
            day_found = False
            for date_obj, req_list in week_data['data'].items():
                 if date_obj == current_day:
                    week_matrix.append(req_list)
                    day_found = True
                    break
            if not day_found:
                week_matrix.append([0] * num_intervals)

        results.append({
            'display': f"Week of {week_start_date.strftime('%Y-%m-%d')} | FTE: {avg_fte:.1f}",
            'week_start_dt': week_start_date,
            'avg_fte': avg_fte,
            'matrix': week_matrix
        })
    return sorted(results, key=lambda x: x['week_start_dt'])

def pareto_analysis(df, value_col, group_by_col, threshold):
    """
    Identifies the top intervals that contribute to a certain percentage
    of the total value (e.g., volume or required staff).
    """
    # Sort by value within each group to find the top contributors
    df_sorted = df.sort_values(by=[group_by_col, value_col], ascending=[True, False])

    df_sorted['CumulativeSum'] = df_sorted.groupby(group_by_col)[value_col].cumsum()

    group_totals = df_sorted.groupby(group_by_col)[value_col].sum().rename('Total')
    df_sorted = df_sorted.join(group_totals, on=group_by_col)

    df_sorted['CumulativePercentage'] = df_sorted['CumulativeSum'] / df_sorted['Total']

    pareto_df = df_sorted[df_sorted['CumulativePercentage'] <= (threshold / 100)].copy()

    pareto_df['Contribution (%)'] = (pareto_df[value_col] / df_sorted['Total']) * 100

    # Sort the final result chronologically by interval for better readability
    pareto_df = pareto_df.sort_values(by='Interval')
    pareto_df['Interval'] = pareto_df['Interval'].apply(lambda t: t.strftime('%H:%M') if isinstance(t, datetime.time) else t)

    return pareto_df[[group_by_col, 'Interval', value_col, 'Contribution (%)', 'CumulativePercentage']]

def download_dataframe_csv(df, filename_prefix):
    """Generates a download button for a DataFrame as CSV."""
    csv = df.to_csv(index=True).encode('utf-8') # index=True for dataframes like scenario_summary where index is meaningful
    st.download_button(
        label="Download Data as CSV",
        data=csv,
        file_name=f"{filename_prefix}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv",
        key=f"download_csv_{filename_prefix}_{id(df)}" # Unique key per df instance
    )

def download_dataframe_csv_no_index(df, filename_prefix):
    """Generates a download button for a DataFrame as CSV (no index)."""
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Download Data as CSV",
        data=csv,
        file_name=f"{filename_prefix}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv",
        key=f"download_csv_{filename_prefix}_{id(df)}" # Unique key per df instance
    )


# ------------------------------------------------------------------------------
#                           TAB 1: STAFFING CALCULATOR
# ------------------------------------------------------------------------------

def calculate_erlang_c_positions(awt, shrinkage, max_occupancy, avg_aht, target, calls):
    """
    Calculates required positions and the resulting KPIs using MultiErlangC.
    This method finds the minimum positions to meet targets and returns the full performance dict.
    """
    interval_seconds = pd.to_timedelta(st.session_state.interval_freq).total_seconds()
    # MultiErlangC expects a grid of parameters to test. We are using it for a single calculation.
    param_grid = {
        "transactions": [calls],
        "aht": [avg_aht],
        "interval": [int(interval_seconds)],
        "asa": [awt],
        "service_level": [target / 100],  # Add service_level to param_grid
        "shrinkage": [shrinkage / 100]
    }
    multi_erlang = MultiErlangC(param_grid=param_grid, n_jobs=1)

    # The required_positions method expects a dictionary of scenarios.
    required_positions_scenarios = {
        "service_level": [target / 100],
        "max_occupancy": [max_occupancy / 100]
    }
    
    # This will return a list of dictionaries, one for each combination of scenarios.
    # Since we have one of each, we get a list with a single dictionary.
    results_list = multi_erlang.required_positions(required_positions_scenarios)
    
    return results_list

def calculate_erlang_c_with_concurrency_positions(awt, shrinkage, max_occupancy, avg_aht, target, calls, concurrency):
    """
    Calculates required positions for a concurrent channel by adjusting the arrival rate.
    """
    if calls == 0:
        # Return a structure that matches the MultiErlangC output for consistency
        return [{'positions': 0, 'service_level': 1, 'occupancy': 0, 'waiting_probability': 0, 'asa': 0}]
    if concurrency <= 0:
        raise ValueError("Concurrency must be greater than 0.")

    adjusted_calls = calls / concurrency
    return calculate_erlang_c_positions(awt, shrinkage, max_occupancy, avg_aht, target, adjusted_calls)

def calculate_transactional_positions(volume, aht, shrinkage, interval_seconds=1800):
    """Calculates required positions for transactional tasks (email, back-office)."""
    if volume == 0 or aht == 0:
        return 0
    workload = volume * aht
    raw_positions = workload / interval_seconds
    return math.ceil(raw_positions / (1 - (shrinkage / 100)))

def calculate_aggregated_kpis(df_slice):
    """
    Calculates volume-weighted KPIs for a given dataframe slice (e.g., a day or a week).
    """
    total_volume = df_slice["Volume"].sum()
    working_hours = st.session_state.get('working_hours', 8.0)  # Default to 8 if not set

    if total_volume == 0:
        return {
            "Total Calls": 0, "Total Raw Positions": 0, "Total Final Positions": 0,
            "Service Level (%)": 100.0, "Occupancy (%)": 0.0,
            "Wait Probability (%)": 0.0, "Overall ASA (s)": 0.0,
            "Required HC per day": 0.0, "Required Raw per day": 0.0
        }

    # Safely get columns using .get() to provide a default value if the column doesn't exist.
    weighted_sl = (df_slice.get('service_level', 0) * df_slice['Volume']).sum() / total_volume
    weighted_occ = (df_slice.get('occupancy', 0) * df_slice['Volume']).sum() / total_volume
    weighted_wp = (df_slice.get('waiting_probability', 0) * df_slice['Volume']).sum() / total_volume
    # Use the 'ASA_s' column for the overall average speed of answer
    weighted_asa = (df_slice.get('ASA_s', 0) * df_slice['Volume']).sum() / total_volume

    raw_positions_sum = df_slice["raw_positions"].sum() if "raw_positions" in df_slice.columns else 0
    
    # Handle final_positions, falling back to Scheduled, and then to 0 if neither exist.
    if "final_positions" in df_slice.columns:
        final_positions_sum = df_slice["final_positions"].sum()
    elif "Scheduled" in df_slice.columns:
        final_positions_sum = df_slice["Scheduled"].sum()
    else:
        final_positions_sum = 0
    
    return {
        "Total Calls": int(total_volume),
        "Total Raw Positions": int(raw_positions_sum),
        "Total Final Positions": int(final_positions_sum),
        "Service Level (%)": weighted_sl * 100,
        "Occupancy (%)": weighted_occ * 100,
        "Wait Probability (%)": weighted_wp * 100,
        "Overall ASA (s)": weighted_asa,
        "Required HC per day": (final_positions_sum / working_hours) if working_hours > 0 else 0,
        "Required Raw per day": (raw_positions_sum / working_hours) if working_hours > 0 else 0
    }

def run_staffing_calculation(params, input_dates_str, day_name_map, week_start_day_name, volume_df):
    """Orchestrator for running staffing calculations for a single channel."""
    staffing_results = []
    intervals_list = st.session_state.intervals
    channel = params['channel_type']
    interval_seconds = (pd.to_timedelta(st.session_state.interval_freq).total_seconds())

    for date_str in input_dates_str:
        current_date = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
        week_start = get_week_start(current_date, week_start_day_name)

        for interval_time, volume in zip(intervals_list, volume_df[date_str]):
            volume_val = validate_and_convert_to_float(volume, "Volume")
            common_data = {
                "Date": pd.to_datetime(date_str), "Day": day_name_map[date_str],
                "Interval": interval_time, "Week_Start_Day": week_start,
                "Volume": volume_val, "AHT": params['aht']
            }
            if volume_val == 0:
                 # Ensure all columns exist even for zero volume intervals
                 staffing_results.append({**common_data, "raw_positions": 0, "final_positions": 0, "service_level": 1.0, "occupancy": 0.0, "waiting_probability": 0.0, "AWT_for_Queued_s": 0.0, "ASA_s": 0.0})
                 continue

            if channel in ["Voice (Erlang-C)", "Chat (Erlang with Concurrency)"]:
                # The MultiErlangC method finds the required positions AND returns the performance KPIs for that number.
                if channel == "Voice (Erlang-C)":
                    kpi_results = calculate_erlang_c_positions(params['awt'], params['shrinkage'], params['max_occupancy'], params['aht'], params['target'], volume_val)
                else:  # Chat
                    kpi_results = calculate_erlang_c_with_concurrency_positions(params['awt'], params['shrinkage'], params['max_occupancy'], params['aht'], params['target'], volume_val, params['concurrency'])

                # Extract all KPIs from the single result dictionary
                kpis = kpi_results[0]
                raw_positions_needed = kpis['positions']

                # Process and store the detailed results
                awt_for_queued = kpis.get('asa', 0)
                wp = kpis.get('waiting_probability', 0)

                result_row = {
                    'raw_positions': raw_positions_needed,
                    'final_positions': math.ceil(raw_positions_needed * (1 / (1 - (params['shrinkage']/100)))),
                    'service_level': kpis.get('service_level', 0),
                    'occupancy': kpis.get('occupancy', 0),
                    'waiting_probability': wp,
                    'AWT_for_Queued_s': awt_for_queued,
                    'ASA_s': awt_for_queued, # Directly use ASA from kpis for consistency
                }
                staffing_results.append({**result_row, **common_data})

            elif channel == "Email / Back Office (Transactional)":
                raw_pos = calculate_transactional_positions(volume_val, params['aht'], params['shrinkage'], interval_seconds)
                transactional_result = {
                    "raw_positions": raw_pos, "final_positions": raw_pos, "service_level": 1.0,
                    "occupancy": 0.0, "waiting_probability": 0.0, "AWT_for_Queued_s": 0.0, "ASA_s": 0.0
                }
                staffing_results.append({**transactional_result, **common_data})

    df = pd.DataFrame(staffing_results)
    if not df.empty:
      df['Week_Start_Day'] = pd.to_datetime(df['Week_Start_Day'])
    return df

# ------------------------------------------------------------------------------
#                       SCHEDULING & COSTING CORE (OR-Tools)
# ------------------------------------------------------------------------------

# --- FIXED: Post-Schedule Performance Calculation ---

# NEW: Interactive Constraint Analyzer Function
def analyze_schedule_feasibility(req_matrix, headcount, schedule_mode, **kwargs):
    """
    Analyzes scheduling constraints against requirements to identify potential
    issues before running the full solver.
    """
    findings = []
    interval_duration_hours = pd.to_timedelta(st.session_state.interval_freq).total_seconds() / 3600
    req_matrix_np = np.array(req_matrix)
    num_days, num_intervals = req_matrix_np.shape
    force_schedule_flag = kwargs.get('force_schedule_insufficient_hc', False)

    # Check 1: Peak Demand vs. Headcount (CRITICAL)
    peak_req = np.max(req_matrix_np)
    if peak_req > headcount:
        if force_schedule_flag:
            findings.append({
                'type': 'INTENTIONAL_UNDERSTAFFING',
                'message': f"Peak requirement of {int(peak_req)} exceeds total headcount of {headcount}. You have enabled scheduling with insufficient headcount.",
                'severity': 'WARNING',
                'suggestion': "The 'Best Fit' model will be run to minimize the gap, but significant understaffing is expected. All other models will be skipped."
            })
        else:
            findings.append({
                'type': 'CAPACITY_EXCEEDED',
                'message': f"Peak requirement of {int(peak_req)} staff exceeds total headcount of {headcount}.",
                'severity': 'CRITICAL',
                'suggestion': f"Increase total headcount to at least {int(peak_req)} or enable the 'Attempt to schedule with insufficient headcount' option in the sidebar to run a 'Best Fit' schedule."
            })

    # Check 2: Total Hours vs. Available Capacity (HIGH)
    total_required_hours = np.sum(req_matrix_np) * interval_duration_hours
    available_work_hours = 0
    if schedule_mode == "Use Pre-defined Shifts":
        shifts_df = kwargs.get('shifts_df')
        work_days_by_shift = kwargs.get('work_days_by_shift')
        if shifts_df is not None and not shifts_df.empty and work_days_by_shift:
            avg_shift_len = shifts_df['Shift Length (hours)'].mean()
            avg_work_days = np.mean(list(work_days_by_shift.values())) if work_days_by_shift else 0
            if avg_work_days > 0 and avg_shift_len > 0:
                 available_work_hours = headcount * avg_work_days * avg_shift_len

    elif schedule_mode == "Optimize Shifts Automatically":
        duration_rules = kwargs.get('duration_rules')
        allowed_durations = kwargs.get('allowed_durations')
        if duration_rules and allowed_durations:
             avg_max_days = np.mean([rules.get('max_days', 5) for rules in duration_rules.values()])
             avg_duration = np.mean(allowed_durations)
             available_work_hours = headcount * avg_max_days * avg_duration

    if available_work_hours > 0 and total_required_hours > (available_work_hours * 0.98):
         findings.append({
            'type': 'INSUFFICIENT_CAPACITY',
            'message': f"Total required hours ({total_required_hours:,.0f}) is very high compared to the estimated available work hours ({available_work_hours:,.0f}). The schedule will be very tight or impossible.",
            'severity': 'HIGH',
            'suggestion': "Increase headcount, allow more flexible work rules (e.g., more days/week), or use longer shifts to increase total capacity."
        })

    # Check 3: Coverage Gap Detection (CRITICAL)
    days_of_week_ordered = kwargs.get('days_of_week_ordered')
    intervals = st.session_state.intervals

    if schedule_mode == "Use Pre-defined Shifts":
        virtual_shifts, _ = expand_shifts_for_solver(kwargs.get('shifts_df'))
        if not virtual_shifts:
            findings.append({
                'type': 'NO_SHIFTS_DEFINED',
                'message': "No valid shifts have been defined.",
                'severity': 'CRITICAL',
                'suggestion': "Go to 'Shift Pattern Definitions' in the sidebar and define at least one valid shift."
            })
        else:
            for d in range(num_days):
                for p in range(num_intervals):
                    if req_matrix_np[d, p] > 0:
                        if not any(vs['availability_coverage'][p] == 1 for vs in virtual_shifts):
                            day_name = days_of_week_ordered[d]
                            interval_time = intervals[p].strftime('%H:%M')
                            findings.append({
                                'type': 'COVERAGE_GAP',
                                'message': f"A requirement at {day_name} {interval_time} cannot be covered. No defined shift is active at this time.",
                                'severity': 'CRITICAL',
                                'suggestion': "Adjust existing shifts to cover this time or add a new shift pattern."
                            })
                            break # Move to next day after finding one gap
                else:
                    continue
                break

    elif schedule_mode == "Optimize Shifts Automatically":
        daily_op_hours = kwargs.get('daily_op_hours')
        for d in range(num_days):
            day_name = days_of_week_ordered[d]
            op_hours = daily_op_hours.get(day_name)
            if not op_hours: continue

            op_start, op_end = op_hours['Start Time'], op_hours['End Time']
            # Special case for 24/7: if start and end are 00:00, it's valid, skip this check.
            if op_start == datetime.time(0, 0) and op_end == datetime.time(0, 0):
                continue

            for p in range(num_intervals):
                if req_matrix_np[d, p] > 0:
                    interval_time = intervals[p]
                    is_in_op_hours = (op_start <= op_end and op_start <= interval_time < op_end) or \
                                     (op_start > op_end and (interval_time >= op_start or interval_time < op_end))
                    if not is_in_op_hours:
                        findings.append({
                            'type': 'REQUIREMENT_OUTSIDE_OPERATING_HOURS',
                            'message': f"A requirement exists on {day_name} at {interval_time.strftime('%H:%M')}, which is outside the defined operational hours for that day.",
                            'severity': 'CRITICAL',
                            'suggestion': "Adjust the operational hours for this day to include this time, or remove the requirement from your forecast."
                        })
                        break
            else:
                continue
            break


    # Check 4: Rule Flexibility (WARNING/CRITICAL)
    if schedule_generation_mode == "Optimize Shifts Automatically":
        duration_rules = kwargs.get('duration_rules')
        for dur, rules in duration_rules.items():
            max_days = rules.get('max_days', 7)
            min_off = rules.get('min_off', 1)
            days_off = 7 - max_days
            if days_off > 0 and days_off < min_off:
                 findings.append({
                    'type': 'RULE_CONFLICT',
                    'message': f"For {dur}hr shifts, working up to {max_days} days leaves {days_off} day(s) off, which is less than the required minimum of {min_off} consecutive off days. This is impossible.",
                    'severity': 'CRITICAL',
                    'suggestion': f"For {dur}hr shifts, reduce the 'Min Consecutive Off' to {days_off} or less, or decrease the maximum 'Work Days/Wk'."
                })
            elif days_off > 0 and days_off == min_off:
                findings.append({
                    'type': 'LOW_FLEXIBILITY',
                    'message': f"For {dur}hr shifts, the work/off day rules ({max_days} on, {min_off} off) leave no flexibility for scheduling off days. This may make finding an optimal schedule difficult.",
                    'severity': 'WARNING',
                    'suggestion': "This is not an error, but it constrains the solver. If it fails, consider reducing the 'Min Consecutive Off' days by one."
                })

    # Return unique findings by message
    unique_findings = []
    seen_messages = set()
    for finding in findings:
        if finding['message'] not in seen_messages:
            unique_findings.append(finding)
            seen_messages.add(finding['message'])
    return unique_findings


def solve_schedule_ortools(required_staff, virtual_shifts, shift_groups, num_employees, work_days_by_shift, objective_type, constraints_config, week_day_names, line_adherence_config=None, force_fit_mode=False):
    """
    Generates a weekly schedule. Enforces that each employee works one shift type
    and adheres to the work-day rules for that specific shift type.
    In force_fit_mode, it relaxes work-life-balance constraints to find a schedule.
    """
    num_days = 7
    num_intervals = len(required_staff[0])
    num_virtual_shifts = len(virtual_shifts)

    if num_employees == 0 or num_virtual_shifts == 0:
        return {'status': 'INFEASIBLE', 'reason': 'No employees or shifts defined.'}

    model = cp_model.CpModel()

    # work[e, vs_id, d]: employee e works virtual_shift vs_id on day d
    work = {}
    for e in range(num_employees):
        for vs_id in range(num_virtual_shifts):
            for d in range(num_days):
                work[e, vs_id, d] = model.NewBoolVar(f'work_{e}_{vs_id}_{d}')

    # An employee can work at most one shift per day.
    for e in range(num_employees):
        for d in range(num_days):
            model.AddAtMostOne(work[e, vs_id, d] for vs_id in range(num_virtual_shifts))

    # --- Constraint: Each employee works only one shift type per week ---
    is_assigned_to_shift = {}
    for e in range(num_employees):
        for shift_name in shift_groups.keys():
            is_assigned_to_shift[e, shift_name] = model.NewBoolVar(f'is_assigned_{e}_{shift_name}')

    # An employee can be assigned to at most one original shift type for the week.
    for e in range(num_employees):
        model.AddAtMostOne(is_assigned_to_shift[e, shift_name] for shift_name in shift_groups.keys())

    # Link the daily work assignments to the weekly shift type assignment.
    for e in range(num_employees):
        for vs in virtual_shifts:
            vs_id = vs['solver_id']
            original_name = vs['original_name']
            for d in range(num_days):
                model.AddImplication(work[e, vs_id, d], is_assigned_to_shift[e, original_name])

    # --- Constraint: Work days per week based on assigned shift type ---
    for e in range(num_employees):
        total_work_days = sum(work[e, vs_id, d] for vs_id in range(num_virtual_shifts) for d in range(num_days))

        # Apply the specific work-day constraint ONLY IF the employee is assigned to that shift type.
        for shift_name, work_days in work_days_by_shift.items():
            if force_fit_mode:
                # RELAXED: Can work UP TO the specified days.
                model.Add(total_work_days <= work_days).OnlyEnforceIf(is_assigned_to_shift[e, shift_name])
            else:
                # STRICT: Must work EXACTLY the specified days.
                model.Add(total_work_days == work_days).OnlyEnforceIf(is_assigned_to_shift[e, shift_name])

        is_assigned_any_shift = model.NewBoolVar(f'is_assigned_any_{e}')
        model.Add(is_assigned_any_shift == sum(is_assigned_to_shift[e, shift_name] for shift_name in shift_groups.keys()))
        model.Add(total_work_days == 0).OnlyEnforceIf(is_assigned_any_shift.Not())


    works_on_day = {}
    for e in range(num_employees):
        for d in range(num_days):
            works_on_day[e, d] = model.NewBoolVar(f'works_on_day_{e}_{d}')
            model.Add(works_on_day[e, d] == sum(work[e, vs_id, d] for vs_id in range(num_virtual_shifts)))

    scheduled_staff = [[model.NewIntVar(0, num_employees, f'sched_{d}_{p}') for p in range(num_intervals)] for d in range(num_days)]
    for d in range(num_days):
        for p in range(num_intervals):
            model.Add(scheduled_staff[d][p] == sum(work[e, vs_id, d] * virtual_shifts[vs_id]['availability_coverage'][p] for e in range(num_employees) for vs_id in range(num_virtual_shifts)))

    # --- RELAXABLE CONSTRAINTS ---
    if not force_fit_mode:
        if constraints_config.get('max_consecutive_work'):
            max_consecutive = constraints_config.get('max_consecutive_work')
            for e in range(num_employees):
                for d in range(num_days): model.Add(sum(works_on_day[e, (d + i) % num_days] for i in range(max_consecutive + 1)) <= max_consecutive)

        min_off = constraints_config.get('min_consecutive_off', 1)
        if min_off > 1:
            for e in range(num_employees):
                for d_start in range(num_days):
                    literals = [works_on_day[e, d_start].Not()]
                    for i in range(1, min_off):
                        literals.append(works_on_day[e, (d_start + i) % num_days])
                    literals.append(works_on_day[e, (d_start + min_off) % num_days].Not())
                    model.AddBoolOr(literals)

    if objective_type == 'meet_or_exceed':
        for d in range(num_days):
            for p in range(num_intervals): model.Add(scheduled_staff[d][p] >= required_staff[d][p])
        total_scheduled_intervals = sum(p for day in scheduled_staff for p in day)
        model.Minimize(total_scheduled_intervals)

    elif objective_type == 'line_adherence' and line_adherence_config:
        target_level = line_adherence_config['target_level']
        target_percent = line_adherence_config['target_percent']
        cap_percent = line_adherence_config['cap_percent']

        capped_scheduled_staff = [[model.NewIntVar(0, num_employees, f'capped_sched_{d}_{p}') for p in range(num_intervals)] for d in range(num_days)]
        for d in range(num_days):
            for p in range(num_intervals):
                if required_staff[d][p] > 0:
                    max_allowed_staff = math.ceil(required_staff[d][p] * cap_percent / 100)
                    model.AddMinEquality(capped_scheduled_staff[d][p], [scheduled_staff[d][p], max_allowed_staff])
                else:
                    model.Add(capped_scheduled_staff[d][p] == 0)

        total_scheduled_intervals = sum(p for day in scheduled_staff for p in day)
        under_target_points = model.NewIntVar(0, 1000000000, 'under_target_points')

        if target_level == 'day':
            daily_shortfalls = []
            for d in range(num_days):
                daily_req = sum(required_staff[d])
                if daily_req > 0:
                    daily_target = int(daily_req * target_percent)
                    daily_contrib = sum(capped_scheduled_staff[d]) * 100
                    shortfall = model.NewIntVar(0, 100000000, f'shortfall_{d}')
                    model.Add(shortfall >= daily_target - daily_contrib)
                    daily_shortfalls.append(shortfall)
            model.Add(under_target_points == sum(daily_shortfalls))

        else: # 'week'
            total_req = sum(v for day in required_staff for v in day)
            if total_req > 0:
                weekly_target = int(total_req * target_percent)
                weekly_contrib = sum(v for day in capped_scheduled_staff for v in day) * 100
                model.Add(under_target_points >= weekly_target - weekly_contrib)

        HUGE_PENALTY = 1000000
        model.Minimize(under_target_points * HUGE_PENALTY + total_scheduled_intervals)

    else: # 'best_fit' model
        over_staff_vars, under_staff_vars = [], []
        for d in range(num_days):
            for p in range(num_intervals):
                over = model.NewIntVar(0, num_employees, f'over_{d}_{p}')
                under = model.NewIntVar(0, num_employees, f'under_{d}_{p}')
                model.Add(scheduled_staff[d][p] - required_staff[d][p] == over - under)
                over_staff_vars.append(over)
                under_staff_vars.append(under)
        objective = (sum(over_staff_vars) * constraints_config.get('overstaff_penalty', 1) + sum(under_staff_vars) * constraints_config.get('understaff_penalty', 10))
        model.Minimize(objective)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 60.0
    # Use all available CPU cores
    # Fallback to 8 if os.cpu_count() is None or 0.
    solver.parameters.num_search_workers = os.cpu_count() or 8
    status = solver.Solve(model)

    results = {'status': solver.StatusName(status)}
    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        roster = []
        for e in range(num_employees):
            emp_row = {'Employee': f'Emp_{e+1}'}
            for d_idx, day_name in enumerate(week_day_names):
                is_off = True
                for vs_id in range(num_virtual_shifts):
                    if solver.Value(work[e, vs_id, d_idx]) == 1:
                        emp_row[day_name] = virtual_shifts[vs_id]['display_name']
                        is_off = False
                        break
                if is_off: emp_row[day_name] = 'OFF'
            roster.append(emp_row)
        results['roster_df'] = pd.DataFrame(roster)
        results['scheduled_staff'] = [[solver.Value(p) for p in day] for day in scheduled_staff]
    return results

@st.cache_data(ttl=3600)
def expand_candidate_shifts(allowed_durations):
    """
    Generates all possible shifts based on allowed durations and 30-min start increments.
    This creates a comprehensive pool of candidates; filtering by operational hours happens in the solver.
    """
    candidate_shifts = []
    solver_id_counter = 0
    num_intervals = 48  # 30-min intervals in 24h
    today = datetime.date.today()

    for start_mins in range(0, 24 * 60, 30):
        start_h, start_m = divmod(start_mins, 60)
        start_time = datetime.time(start_h, start_m)

        for length_hours in allowed_durations:
            start_interval = start_h * 2 + start_m // 30
            length_intervals = int(length_hours * 2)

            coverage = [0] * num_intervals
            for i in range(length_intervals):
                coverage[(start_interval + i) % num_intervals] = 1

            candidate_shifts.append({
                'solver_id': solver_id_counter,
                'start_time': start_time,
                'length_hours': length_hours,
                'coverage': coverage
            })
            solver_id_counter += 1
    return candidate_shifts

def solve_schedule_with_shift_optimization(required_staff, week_day_names, optimization_config):
    """
    Finds the optimal set of shifts and generates a roster to meet demand based on detailed user constraints,
    including day-specific operational hours.
    In force_fit mode, it relaxes work-life-balance constraints to find a schedule.
    """
    total_employees = optimization_config['total_headcount']
    if total_employees == 0:
        return {'status': 'INFEASIBLE', 'reason': 'Total headcount for optimization is zero.'}

    force_fit_mode = optimization_config.get('force_fit', False)
    num_days = 7
    num_intervals = 48
    model = cp_model.CpModel()

    # --- 1. Pre-generate candidate shifts based on user rules ---
    candidate_shifts = expand_candidate_shifts(
        optimization_config['allowed_durations']
    )
    if not candidate_shifts:
        return {'status': 'INFEASIBLE', 'reason': 'No possible shifts can be created with the given duration rules.'}
    num_candidate_shifts = len(candidate_shifts)

    # --- 2. Define decision variables ---
    shift_is_chosen = [model.NewBoolVar(f'chosen_{i}') for i in range(num_candidate_shifts)]
    work = {}
    for e in range(total_employees):
        for cs_id in range(num_candidate_shifts):
            for d in range(num_days):
                work[e, cs_id, d] = model.NewBoolVar(f'work_{e}_{cs_id}_{d}')

    employee_assigned_to_duration = {}
    for e in range(total_employees):
        for dur in optimization_config['allowed_durations']:
            employee_assigned_to_duration[e, dur] = model.NewBoolVar(f'emp_{e}_is_{dur}hr')

    # --- 3. Define constraints ---
    model.Add(sum(shift_is_chosen) <= optimization_config['max_unique_shifts'])

    # --- Constraint: Shift must be within daily operational hours ---
    for cs_id, shift in enumerate(candidate_shifts):
        for d_idx, day_name in enumerate(week_day_names):
            op_hours = optimization_config['daily_op_hours'][day_name]
            op_start, op_end = op_hours['Start Time'], op_hours['End Time']

            # --- NEW: 24/7 Operation Handling ---
            # If start and end are both 00:00, it's a 24/7 operation. All shifts are valid for this day.
            if op_start == datetime.time(0, 0) and op_end == datetime.time(0, 0):
                continue  # Skip to the next day, as any shift is valid.

            # --- ENHANCED LOGIC FOR OVERNIGHT WINDOWS ---
            shift_start_dt = datetime.datetime.combine(datetime.date.today(), shift['start_time'])
            shift_end_dt = shift_start_dt + datetime.timedelta(hours=shift['length_hours'])

            op_start_dt = datetime.datetime.combine(datetime.date.today(), op_start)
            op_end_dt = datetime.datetime.combine(datetime.date.today(), op_end)
            
            # If the operational window crosses midnight (e.g., 22:00 to 06:00), add a day to the end time for correct comparison.
            if op_end_dt <= op_start_dt:
                op_end_dt += datetime.timedelta(days=1)

            # A shift is valid if its start and end datetimes fall within the operational datetimes.
            is_valid = (shift_start_dt >= op_start_dt) and (shift_end_dt <= op_end_dt)

            if not is_valid:
                # If the shift is NOT valid for this day, no employee can be assigned to it.
                for e in range(total_employees):
                    model.Add(work[e, cs_id, d_idx] == 0)

    # --- Shift Consistency Logic ---
    if optimization_config.get('shift_consistency', False):
        employee_assigned_to_shift_type = {}
        for e in range(total_employees):
            for cs_id in range(num_candidate_shifts):
                employee_assigned_to_shift_type[e, cs_id] = model.NewBoolVar(f'emp_{e}_is_shift_{cs_id}')

            # Constraint: At most one shift type per employee for the week.
            model.AddAtMostOne(employee_assigned_to_shift_type[e, cs_id] for cs_id in range(num_candidate_shifts))

            # Link daily work to the weekly shift assignment.
            for cs_id in range(num_candidate_shifts):
                for d in range(num_days):
                    model.AddImplication(work[e, cs_id, d], employee_assigned_to_shift_type[e, cs_id])

    for e in range(total_employees):
        model.AddExactlyOne(employee_assigned_to_duration[e, dur] for dur in optimization_config['allowed_durations'])
        for d in range(num_days):
            model.AddAtMostOne(work[e, cs_id, d] for cs_id in range(num_candidate_shifts))

        for cs_id, shift in enumerate(candidate_shifts):
            for d in range(num_days):
                model.AddImplication(work[e, cs_id, d], employee_assigned_to_duration[e, shift['length_hours']])
                model.AddImplication(work[e, cs_id, d], shift_is_chosen[cs_id])

        total_work_days = sum(work[e, cs_id, d] for cs_id in range(num_candidate_shifts) for d in range(num_days))
        works_on_day = [model.NewBoolVar(f'works_day_{e}_{d}') for d in range(num_days)]
        for d in range(num_days):
            model.Add(works_on_day[d] == sum(work[e, cs_id, d] for cs_id in range(num_candidate_shifts)))

        for dur, rules in optimization_config['duration_rules'].items():
            if force_fit_mode:
                # RELAXED: Only enforce the maximum number of work days.
                model.Add(total_work_days <= rules.get('max_days', 7)).OnlyEnforceIf(employee_assigned_to_duration[e, dur])
            else:
                # STRICT: Enforce both min and max work days.
                model.AddLinearConstraint(
                    total_work_days,
                    rules.get('min_days', 1),
                    rules.get('max_days', 7)
                ).OnlyEnforceIf(employee_assigned_to_duration[e, dur])

            if not force_fit_mode:
                min_off = rules.get('min_off', 1)
                if min_off > 1:
                    for d_start in range(num_days):
                        literals = [works_on_day[d_start].Not()]
                        for i in range(1, min_off):
                            literals.append(works_on_day[(d_start + i) % num_days])
                        literals.append(works_on_day[(d_start + min_off) % num_days].Not())
                        model.AddBoolOr(literals).OnlyEnforceIf(employee_assigned_to_duration[e, dur])

    # --- Staffing & Distribution Constraints ---
    for cs_id in range(num_candidate_shifts):
        total_assignments = model.NewIntVar(0, total_employees * num_days, f'total_assign_{cs_id}')
        model.Add(total_assignments == sum(work[e, cs_id, d] for e in range(total_employees) for d in range(num_days)))
        model.Add(total_assignments >= optimization_config['min_agents_per_shift']).OnlyEnforceIf(shift_is_chosen[cs_id])
        model.Add(total_assignments <= optimization_config['max_agents_per_shift']).OnlyEnforceIf(shift_is_chosen[cs_id])

    # --- Dynamic Distribution Cap per Duration ---
    total_all_assignments = model.NewIntVar(0, total_employees * num_days, 'all_assign')
    model.Add(total_all_assignments == sum(work[e, cs_id, d] for e in range(total_employees) for cs_id in range(num_candidate_shifts) for d in range(num_days)))

    for dur, cap_rules in optimization_config['distribution_caps'].items():
        if cap_rules.get('enabled', False):
            dur_shift_ids = [s['solver_id'] for s in candidate_shifts if s['length_hours'] == dur]
            if dur_shift_ids:
                total_dur_assignments = model.NewIntVar(0, total_employees * num_days, f'dur_{dur}_assign')
                model.Add(total_dur_assignments == sum(work[e, cs_id, d] for e in range(total_employees) for cs_id in dur_shift_ids for d in range(num_days)))
                model.Add(100 * total_dur_assignments <= cap_rules['percent'] * total_all_assignments)

    # --- 4. Link to requirements and set objective ---
    scheduled_staff = [[model.NewIntVar(0, total_employees, f'sched_{d}_{p}') for p in range(num_intervals)] for d in range(num_days)]
    for d in range(num_days):
        for p in range(num_intervals):
            model.Add(scheduled_staff[d][p] == sum(work[e, cs_id, d] * candidate_shifts[cs_id]['coverage'][p] for e in range(total_employees) for cs_id in range(num_candidate_shifts)))

    # --- CHOOSE OBJECTIVE BASED ON MODEL TYPE ---
    model_type = optimization_config.get('model_type', 'best_fit')

    if model_type == 'line_adherence':
        cap_percent = optimization_config.get('cap_percent', 105)
        target_percent = optimization_config.get('target_percent', 95)
        target_level = optimization_config.get('target_level', 'day')

        capped_scheduled_staff = [[model.NewIntVar(0, total_employees, f'capped_sched_{d}_{p}') for p in range(num_intervals)] for d in range(num_days)]
        for d in range(num_days):
            for p in range(num_intervals):
                if required_staff[d][p] > 0:
                    max_allowed_staff = math.ceil(required_staff[d][p] * cap_percent / 100)
                    model.AddMinEquality(capped_scheduled_staff[d][p], [scheduled_staff[d][p], max_allowed_staff])
                else:
                    model.Add(capped_scheduled_staff[d][p] == 0)

        if target_level == 'day':
            for d in range(num_days):
                daily_req = sum(required_staff[d])
                if daily_req > 0:
                    daily_target_contribution = daily_req * target_percent
                    daily_contribution = sum(capped_scheduled_staff[d])
                    model.Add(daily_contribution * 100 >= daily_target_contribution)
        else:  # 'week'
            total_req = sum(v for day in required_staff for v in day)
            if total_req > 0:
                weekly_target_contribution = total_req * target_percent
                weekly_contribution = sum(v for day in capped_scheduled_staff for v in day)
                model.Add(weekly_contribution * 100 >= weekly_target_contribution)

        total_scheduled_intervals = sum(p for day in scheduled_staff for p in day)
        model.Minimize(total_scheduled_intervals)

    else:  # Default to 'best_fit'
        over_staff = [[model.NewIntVar(0, total_employees, f'over_{d}_{p}') for p in range(num_intervals)] for d in range(num_days)]
        under_staff = [[model.NewIntVar(0, total_employees * 2, f'under_{d}_{p}') for p in range(num_intervals)] for d in range(num_days)]
        for d in range(num_days):
            for p in range(num_intervals):
                model.Add(scheduled_staff[d][p] - required_staff[d][p] == over_staff[d][p] - under_staff[d][p])

        overstaffing_cost = model.NewIntVar(0, total_employees * num_days * num_intervals * 10, 'over_cost')
        understaffing_cost = model.NewIntVar(0, sum(sum(day) for day in required_staff) * total_employees * 10, 'under_cost')

        model.Add(overstaffing_cost == sum(over_staff[d][p] for d in range(num_days) for p in range(num_intervals)))
        model.Add(understaffing_cost == sum(under_staff[d][p] * (required_staff[d][p] + 1) for d in range(num_days) for p in range(num_intervals)))
        model.Minimize(overstaffing_cost * optimization_config['overstaff_penalty'] + understaffing_cost * optimization_config['understaff_penalty'])

    # --- 5. Solve and post-process ---
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 120.0
    solver.parameters.num_search_workers = os.cpu_count() or 8
    status = solver.Solve(model)

    results = {'status': solver.StatusName(status)}
    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        virtual_shifts_generated, chosen_shift_map = [], {}
        final_shift_id_counter = 1
        for cs_id in range(num_candidate_shifts):
            if solver.Value(shift_is_chosen[cs_id]):
                candidate = candidate_shifts[cs_id]
                end_time = (datetime.datetime.combine(datetime.date.today(), candidate['start_time']) + datetime.timedelta(hours=candidate['length_hours'])).time()
                new_name = f"Opti-Shift {final_shift_id_counter}: {candidate['start_time'].strftime('%H:%M')}-{end_time.strftime('%H:%M')} ({candidate['length_hours']}hr)"

                virtual_shifts_generated.append({
                    'display_name': new_name, 'original_name': new_name,
                    'availability_coverage': candidate['coverage'], 'payable_coverage': candidate['coverage']
                })
                chosen_shift_map[cs_id] = new_name
                final_shift_id_counter += 1

        roster = []
        for e in range(total_employees):
            emp_row = {'Employee': f'Emp_{e+1}'}
            for d_idx, day_name in enumerate(week_day_names):
                is_off = True
                for cs_id in chosen_shift_map.keys():
                    if solver.Value(work[e, cs_id, d_idx]) == 1:
                        emp_row[day_name] = chosen_shift_map[cs_id]
                        is_off = False; break
                if is_off: emp_row[day_name] = 'OFF'
            roster.append(emp_row)

        results['roster_df'] = pd.DataFrame(roster)
        results['scheduled_staff'] = [[solver.Value(p) for p in day] for day in scheduled_staff]
        results['virtual_shifts_generated'] = virtual_shifts_generated
    return results

def calculate_schedule_cost(roster_df, virtual_shifts, cost_config, week_start_dt, day_order):
    """
    Calculates the detailed cost of a generated schedule using vectorized Pandas operations.
    """
    if roster_df.empty or not virtual_shifts:
        return {'Base Pay': 0, 'Overtime Premium': 0, 'Shift Differentials': 0, 'Day/Holiday Premiums': 0, 'Total': 0}, pd.DataFrame(), pd.DataFrame()

    # --- Configuration ---
    base_rate = cost_config['base_rate']
    ot_threshold = cost_config['ot_threshold']
    ot_multiplier = cost_config['ot_multiplier']
    shift_diffs = cost_config['shift_differentials']
    day_diffs = cost_config['day_differentials']
    holidays = cost_config['holidays']
    num_intervals = 48
    interval_duration_hours = pd.to_timedelta(st.session_state.interval_freq).total_seconds() / 3600
    intervals_times = st.session_state.intervals

    # --- 1. Prepare DataFrames for Merging ---

    # Work Log: Melt roster to long format
    work_log = roster_df.melt(id_vars=['Employee'], var_name='Day', value_name='Shift').rename(columns={'Shift': 'display_name'})
    work_log = work_log[work_log['display_name'] != 'OFF']

    if work_log.empty:
         return {'Base Pay': 0, 'Overtime Premium': 0, 'Shift Differentials': 0, 'Day/Holiday Premiums': 0, 'Total': 0}, pd.DataFrame(), pd.DataFrame()

    # Shift Coverage: Expand virtual shifts into payable intervals
    shift_coverage_list = []
    for shift in virtual_shifts:
        for i, is_payable in enumerate(shift['payable_coverage']):
            if is_payable:
                shift_coverage_list.append({'display_name': shift['display_name'], 'Interval_idx': i})
    shift_coverage_df = pd.DataFrame(shift_coverage_list)

    # Day Premiums: Create a map for days and holidays
    day_premium_list = []
    for i, day_name in enumerate(day_order):
        date = week_start_dt + datetime.timedelta(days=i)
        date_str = date.strftime('%Y-%m-%d')
        prem = {'mult': 1.0, 'add': 0.0}
        if date_str in holidays:
            rule = holidays[date_str]
            if rule['type'] == 'Multiplier': prem['mult'] = rule.get('value', 1.0)
            else: prem['add'] = rule.get('value', 0.0)
        elif day_name in day_diffs:
            rule = day_diffs[day_name]
            if rule['type'] == 'Multiplier': prem['mult'] = rule.get('value', 1.0)
            else: prem['add'] = rule.get('value', 0.0)
        day_premium_list.append({'Day': day_name, 'day_prem_mult': prem['mult'], 'day_prem_add': prem['add']})
    day_premium_df = pd.DataFrame(day_premium_list)

    # Shift/Time Premiums: Create a map for intervals
    interval_premium_map = [{'name': None, 'type': None, 'value': 0.0}] * num_intervals
    for i, interval_time in enumerate(intervals_times):
        for _, diff in shift_diffs.iterrows():
            if pd.notna(diff['Start Time']) and pd.notna(diff['End Time']):
                start_t, end_t = diff['Start Time'], diff['End Time']
                is_premium = (start_t <= end_t and start_t <= interval_time < end_t) or \
                             (start_t > end_t and (interval_time >= start_t or interval_time < end_t))
                if is_premium:
                    interval_premium_map[i] = {'name': diff['Name'], 'type': diff['Premium Type'], 'value': diff['Premium']}
    interval_premium_df = pd.DataFrame(interval_premium_map).reset_index().rename(columns={'index': 'Interval_idx'})

    # --- 2. Merge and Create Master Cost DataFrame ---
    # Merge work log with shift interval details
    cost_df = pd.merge(work_log, shift_coverage_df, on='display_name')
    # Merge with day premiums
    cost_df = pd.merge(cost_df, day_premium_df, on='Day')
    # Merge with interval-based shift premiums
    cost_df = pd.merge(cost_df, interval_premium_df, on='Interval_idx')

    # --- 3. Vectorized Calculations ---

    # Base Pay
    cost_df['Base_Pay'] = base_rate * interval_duration_hours

    # OT Calculation
    # Use cumcount to track worked intervals per employee for OT threshold
    cost_df = cost_df.sort_values(['Employee', 'Day', 'Interval_idx'])
    cost_df['cumulative_hours'] = (cost_df.groupby('Employee').cumcount() + 1) * interval_duration_hours
    cost_df['Is_OT'] = cost_df['cumulative_hours'] > ot_threshold
    cost_df['OT_Premium'] = np.where(cost_df['Is_OT'], cost_df['Base_Pay'] * (ot_multiplier - 1), 0)

    # Day/Holiday Premiums
    cost_df['Day_Premium'] = (cost_df['Base_Pay'] * (cost_df['day_prem_mult'] - 1)) + (cost_df['day_prem_add'] * interval_duration_hours)

    # Shift Premiums
    is_percentage = cost_df['type'] == 'Percentage'
    is_additive = cost_df['type'] == 'Additive ($)'
    cost_df['Shift_Premium'] = np.where(is_percentage, cost_df['Base_Pay'] * (cost_df['value'] / 100), 0)
    cost_df['Shift_Premium'] += np.where(is_additive, cost_df['value'] * interval_duration_hours, 0)

    # Total Pay
    cost_df['Total_Pay'] = cost_df['Base_Pay'] + cost_df['OT_Premium'] + cost_df['Day_Premium'] + cost_df['Shift_Premium']

    # --- 4. Final Aggregation and Formatting ---

    # Total Cost Breakdown
    total_cost_breakdown = {
        'Base Pay': cost_df['Base_Pay'].sum(),
        'Overtime Premium': cost_df['OT_Premium'].sum(),
        'Shift Differentials': cost_df['Shift_Premium'].sum(),
        'Day/Holiday Premiums': cost_df['Day_Premium'].sum(),
    }
    total_cost_breakdown['Total'] = sum(total_cost_breakdown.values())

    # Employee Weekly Hours
    weekly_hours_df = cost_df.groupby('Employee')['cumulative_hours'].max().reset_index()
    weekly_hours_df.rename(columns={'cumulative_hours': 'Weekly Hours'}, inplace=True)

    # Format final details DataFrame for display
    cost_details_final = cost_df.copy()
    cost_details_final['Interval'] = cost_details_final['Interval_idx'].apply(lambda i: intervals_times[i].strftime('%H:%M'))
    cost_details_final.rename(columns={'display_name': 'Shift'}, inplace=True)

    # Select and reorder columns
    final_cols = ['Employee', 'Day', 'Interval', 'Shift', 'Is_OT', 'Base_Pay', 'OT_Premium', 'Shift_Premium', 'Day_Premium', 'Total_Pay']
    cost_details_final = cost_details_final[final_cols]

    return total_cost_breakdown, cost_details_final, weekly_hours_df


# ------------------------------------------------------------------------------
#                           VISUALIZATION & DISPLAY FUNCTIONS
# ------------------------------------------------------------------------------
def display_comprehensive_results(solution_data, cost_data, requirements_data, key_prefix, day_order, virtual_shifts):
    """Displays the full dashboard for a single schedule solution."""
    roster_df = solution_data['roster_df']
    cost_breakdown, cost_details_df, weekly_hours_df = cost_data

    # --- DEFENSIVE CHECK FOR DATA STRUCTURE ---
    if isinstance(requirements_data, dict) and 'base' in requirements_data and 'inflated' in requirements_data:
        # New, correct data structure
        base_req_matrix = np.array(requirements_data['base'])
        inflated_req_matrix = np.array(requirements_data['inflated'])
        has_base_req = True
    else:
        # Fallback for old data structure or error
        st.warning("Could not find base requirement data. Displaying inflated requirement only. Please re-run the schedule generation on Tab 2 to see the full comparison.", icon="⚠️")
        inflated_req_matrix = np.array(requirements_data)  # Assumes it's the list of inflated reqs
        base_req_matrix = inflated_req_matrix  # Set base to inflated to prevent crashes
        has_base_req = False

    sched_matrix = np.array(solution_data['scheduled_staff'])

    # Calculate metrics against the INFLATED requirement, as this was the solver's target
    diff_matrix = sched_matrix - inflated_req_matrix
    total_required_intervals = np.sum(inflated_req_matrix)
    total_scheduled_intervals = np.sum(sched_matrix)
    overstaffing = np.sum(diff_matrix[diff_matrix > 0])
    understaffing = -np.sum(diff_matrix[diff_matrix < 0])
    coverage_met = (total_scheduled_intervals - overstaffing) / total_required_intervals if total_required_intervals > 0 else 1.0

    interval_duration_hours = pd.to_timedelta(st.session_state.interval_freq).total_seconds() / 3600
    vto_hours = overstaffing * interval_duration_hours
    ot_needed_hours = understaffing * interval_duration_hours

    fte_metrics = calculate_fte_metrics_from_matrix(
        inflated_req_matrix.tolist(),
        st.session_state.get('working_hours', 8.0),
        st.session_state.get('working_days', 5.0)
    )

    st.subheader("Performance & Cost Summary")
    # Display KPIs as a DataFrame for a cleaner look
    kpi_data = {
        'Metric': [
            'Total Labor Cost',
            'Coverage Met',
            'Understaffed Intervals',
            'Overstaffed Intervals',
            'VTO Opportunities (Hours)',
            'OT Needed (Hours)',
            'Inflated HC (Avg)',
            'Inflated HC (Peak Day)'
        ],
        'Value': [
            f"${cost_breakdown['Total']:,.2f}",
            f"{coverage_met:.2%}",
            f"{understaffing:,.0f}",
            f"{overstaffing:,.0f}",
            f"{vto_hours:,.1f}",
            f"{ot_needed_hours:,.1f}",
            f"{fte_metrics['avg_fte']:.2f}",
            f"{fte_metrics['peak_fte']:.2f}"
        ]
    }
    kpi_df = pd.DataFrame(kpi_data).set_index('Metric')
    st.dataframe(kpi_df, use_container_width=True)
    download_dataframe_csv(kpi_df, f"{key_prefix}_kpi_summary")


    st.markdown("---")
    st.markdown("##### Detailed Cost Breakdown")
    cost_df = pd.DataFrame.from_dict(cost_breakdown, orient='index', columns=['Amount'])
    st.dataframe(
        cost_df.style.format('${:,.2f}'),
        use_container_width=True
    )
    download_dataframe_csv(cost_df, f"{key_prefix}_cost_breakdown")

    # NEW PLOT: Daily Cost Breakdown
    if not cost_details_df.empty:
        st.markdown("##### Daily Cost Breakdown Chart")
        daily_costs_summary = cost_details_df.groupby('Day').agg(
            Base_Pay=('Base_Pay', 'sum'),
            OT_Premium=('OT_Premium', 'sum'),
            Shift_Premium=('Shift_Premium', 'sum'),
            Day_Premium=('Day_Premium', 'sum')
        ).reindex(day_order).fillna(0)

        fig_daily_cost = go.Figure()
        fig_daily_cost.add_trace(go.Bar(x=daily_costs_summary.index, y=daily_costs_summary['Base_Pay'], name='Base Pay', marker_color='#1f77b4'))
        fig_daily_cost.add_trace(go.Bar(x=daily_costs_summary.index, y=daily_costs_summary['OT_Premium'], name='OT Premium', marker_color='#ff7f0e'))
        fig_daily_cost.add_trace(go.Bar(x=daily_costs_summary.index, y=daily_costs_summary['Shift_Premium'], name='Shift Premium', marker_color='#2ca02c'))
        fig_daily_cost.add_trace(go.Bar(x=daily_costs_summary.index, y=daily_costs_summary['Day_Premium'], name='Day/Holiday Premium', marker_color='#d62728'))

        fig_daily_cost.update_layout(
            barmode='stack',
            title='Daily Cost Breakdown',
            xaxis_title='Day of Week',
            yaxis_title='Cost ($)',
            legend_title='Cost Component'
        )
        st.plotly_chart(fig_daily_cost, use_container_width=True, key=f"{key_prefix}_daily_cost_chart")


    st.subheader("Generated Roster & Adherence")
    # --- ADDED 'Performance Simulation' Tab ---
    tab_roster, tab_summary, tab_adherence = st.tabs(["Roster Table", "📊 Roster Summary", "📈 Line Adherence Analysis"])
    with tab_roster:
        st.dataframe(roster_df.set_index('Employee'))
        download_dataframe_csv(roster_df.set_index('Employee'), f"{key_prefix}_roster")

    with tab_summary:
        st.markdown("##### Staffing Grid (Employees per Shift per Day)")

        roster_long = roster_df.melt(id_vars=['Employee'], var_name='Day', value_name='Shift')
        roster_working = roster_long[roster_long['Shift'] != 'OFF'].copy()

        if not roster_working.empty:
            staffing_grid = pd.crosstab(
                index=roster_working['Shift'],
                columns=roster_working['Day']
            )
            staffing_grid = staffing_grid.reindex(columns=day_order, fill_value=0)
            st.dataframe(staffing_grid)
            download_dataframe_csv(staffing_grid, f"{key_prefix}_staffing_grid")
        else:
            st.info("No employees were scheduled to work.")

        st.markdown("---")
        st.markdown("##### Shift Coverage Breakdown Chart")
        st.info("This chart shows how many staff each shift type contributes to the scheduled total per interval for a selected day.")

        # NEW PLOT: Shift Contribution Stacked Bar Chart
        if not roster_working.empty and virtual_shifts:
            intervals_str = [t.strftime('%H:%M') for t in st.session_state.intervals]

            # Calculate coverage for each shift type dynamically
            shift_coverage_by_interval = {}
            # Get all unique shifts from the original definition to ensure all possible shifts are considered
            all_defined_shifts_names = [vs['display_name'] for vs in virtual_shifts]

            for original_shift_name in sorted(list(set(roster_working['Shift'].tolist() + all_defined_shifts_names))):
                if original_shift_name == 'OFF': continue

                temp_shift_coverage = {day: [0] * len(intervals_str) for day in day_order}

                # Find the virtual shift that matches this display name
                vs_found = next((vs for vs in virtual_shifts if vs['display_name'] == original_shift_name), None)

                if vs_found: # Only process if the shift definition exists
                    for _, emp_row in roster_df.iterrows():
                        for d_idx, day_name in enumerate(day_order):
                            if emp_row[day_name] == original_shift_name:
                                for p_idx, is_available in enumerate(vs_found['availability_coverage']):
                                    if is_available == 1:
                                        temp_shift_coverage[day_name][p_idx] += 1

                    shift_coverage_by_interval[original_shift_name] = temp_shift_coverage


            selected_day_coverage = st.selectbox("Select a day for shift coverage analysis:", options=day_order, key=f"{key_prefix}_shift_coverage_day_select")

            fig_shift_coverage = go.Figure()

            # Sort shifts by the total staff they contribute on the selected day for consistent order
            sorted_shifts_for_plot = sorted(
                shift_coverage_by_interval.keys(),
                key=lambda s: sum(shift_coverage_by_interval[s].get(selected_day_coverage, [0])),
                reverse=True
            )

            for shift_name in sorted_shifts_for_plot:
                day_data = shift_coverage_by_interval[shift_name]
                if selected_day_coverage in day_data:
                    fig_shift_coverage.add_trace(go.Bar(
                        x=intervals_str,
                        y=day_data[selected_day_coverage],
                        name=shift_name,
                        hovertemplate="Interval: %{x}<br>Shift: %{full_data.name}<br>Staff: %{y}<extra></extra>"
                    ))

            fig_shift_coverage.update_layout(
                barmode='stack',
                title=f"Shift Contribution to Scheduled Staff for {selected_day_coverage}",
                xaxis_title="Time Interval",
                yaxis_title="Scheduled Staff",
                legend_title="Shift Type"
            )
            st.plotly_chart(fig_shift_coverage, use_container_width=True, key=f"{key_prefix}_shift_coverage_chart")
        else:
            st.info("No scheduled staff data to display shift coverage.")

        st.markdown("---")
        st.markdown("##### Daily Staffing Details")

        if not roster_working.empty:
            daily_details_rows = []
            for day in day_order:
                day_roster = roster_working[roster_working['Day'] == day]
                row = {'Day': day, 'Total Staffed': len(day_roster)}

                # Group by shift and format the employee list
                shift_groups = day_roster.groupby('Shift')
                for shift_name, group in shift_groups:
                    # Sort employees numerically by extracting the number from 'Emp_X'
                    try:
                        employees = sorted(group['Employee'].tolist(), key=lambda x: int(re.search(r'\d+', x).group()))
                    except (AttributeError, ValueError):
                        employees = sorted(group['Employee'].tolist()) # Fallback to alphabetical sort

                    emp_list = ", ".join(employees)
                    row[shift_name] = f"({len(group)}): {emp_list}"

                daily_details_rows.append(row)

            if daily_details_rows:
                details_df = pd.DataFrame(daily_details_rows).set_index('Day')

                # Ensure all shift columns exist and are ordered
                all_shift_names = sorted(roster_working['Shift'].unique())
                final_cols = ['Total Staffed']
                for shift in all_shift_names:
                    if shift not in details_df.columns:
                        details_df[shift] = ''
                    final_cols.append(shift)

                details_df = details_df.reindex(day_order) # Ensure correct day order
                details_df = details_df[final_cols]
                details_df = details_df.fillna('') # Fill NaN for shifts not worked on a particular day

                st.dataframe(details_df, use_container_width=True)
                download_dataframe_csv(details_df, f"{key_prefix}_daily_staffing_details")
        else:
            st.info("No staff were assigned to any shifts for this week.")

        st.markdown("---")
        st.markdown("##### Employee Weekly Hours Distribution")
        st.info("This chart shows the total hours worked by each employee over the week. Useful for balancing workload.")
        # NEW PLOT: Employee Weekly Hours Distribution
        if not weekly_hours_df.empty:
            fig_hours_dist = go.Figure(data=[go.Bar(x=weekly_hours_df['Employee'], y=weekly_hours_df['Weekly Hours'])])
            fig_hours_dist.update_layout(
                title='Employee Weekly Hours Distribution',
                xaxis_title='Employee',
                yaxis_title='Total Weekly Hours',
                xaxis={'categoryorder':'total ascending'}
            )
            st.plotly_chart(fig_hours_dist, use_container_width=True, key=f"{key_prefix}_hours_dist_chart")
            download_dataframe_csv(weekly_hours_df, f"{key_prefix}_weekly_hours_distribution")
        else:
            st.info("No employee hours data to display.")


    with tab_adherence:
        adherence_cap_percent = 105
        target_adherence_percent = 95
        if solution_data.get('config') and solution_data['config'].get('cap_percent') is not None:
            adherence_cap_percent = solution_data['config']['cap_percent']
        if solution_data.get('config') and solution_data['config'].get('target_percent') is not None:
            target_adherence_percent = solution_data['config']['target_percent']

        adherence_metrics = calculate_adherence_metrics(inflated_req_matrix, sched_matrix, adherence_cap_percent, day_order)
        weekly_adherence = adherence_metrics['weekly_adherence']
        daily_adherence = adherence_metrics['daily_adherence']
        adherence_df = adherence_metrics['adherence_df']

        st.metric(
            label=f"**Weekly Weighted & Capped Line Adherence (vs. Inflated Req, Target: {target_adherence_percent}%)**",
            value=f"{weekly_adherence:.2f}%",
            delta=f"{weekly_adherence - target_adherence_percent:.2f}% vs Target"
        )
        st.markdown("---")

        st.markdown("**Daily Weighted & Capped Line Adherence**")
        daily_cols = st.columns(7)
        for i, day_name in enumerate(day_order):
            daily_cols[i].metric(label=day_name, value=f"{daily_adherence.get(day_name, 0.0):.1f}%")

        st.markdown("---")

        tab_chart_adherence, tab_table_adherence = st.tabs(["📊 Interval Adherence Chart", "📋 Adherence Data Table"])

        with tab_chart_adherence:
            st.markdown("**Interval-Level Capped Adherence**")
            selected_day = st.selectbox("Select a day to analyze:", options=day_order, key=f"{key_prefix}_day_select_adherence")

            day_df = adherence_df[adherence_df['Day'] == selected_day]

            y_values = day_df['Capped Adherence (%)']
            colors = ['#2ca02c' if x >= 100 else ('#ff7f0e' if x > 0 else '#d62728') for x in y_values]

            fig_adherence = go.Figure()
            fig_adherence.add_trace(go.Bar(
                x=day_df['Interval'],
                y=y_values,
                marker_color=colors,
                name='Adherence',
                hovertemplate=(
                    "<b>%{x}</b><br>"
                    "Inflated Required: %{customdata[0]}<br>"
                    "Scheduled: %{customdata[1]}<br>"
                    "Raw Adherence: %{customdata[2]:.1f}%<br>"
                    "<b>Capped Adherence: %{y:.1f}%</b><extra></extra>"
                ),
                customdata=day_df[['Required', 'Scheduled', 'Raw Adherence (%)']].values
            ))

            if not day_df.empty:
                fig_adherence.add_shape(type="line", x0=day_df['Interval'].iloc[0], y0=100, x1=day_df['Interval'].iloc[-1], y1=100,
                                        line=dict(color="black", width=2, dash="dash"), name="100% Target")

            fig_adherence.update_layout(
                title=f"Line Adherence for {selected_day} (Capped at {adherence_cap_percent}%)",
                xaxis_title="Time Interval",
                yaxis_title="Capped Adherence %",
                yaxis_range=[0, adherence_cap_percent * 1.1],
                showlegend=False
            )
            st.plotly_chart(fig_adherence, use_container_width=True, key=f"{key_prefix}_adherence_bar_chart")

        with tab_table_adherence:
            st.markdown("**Detailed Adherence Calculation Data**")
            st.info("This table shows how adherence is calculated against the inflated requirement, including capping, aligning with the solver's logic.")
            st.dataframe(adherence_df.style.format({
                'Raw Adherence (%)': '{:.2f}%',
                'Capped Adherence (%)': '{:.2f}%',
                'Capped Scheduled Contribution': '{:.2f}'
            }).background_gradient(
                cmap='RdYlGn',
                subset=['Raw Adherence (%)'],
                vmin=0,
                vmax=adherence_cap_percent
            ), use_container_width=True, height=500)
            download_dataframe_csv(adherence_df, f"{key_prefix}_adherence_details")


    st.subheader("Schedule vs. Requirement Deep Dive")
    intervals_str = [t.strftime('%H:%M') for t in st.session_state.intervals]

    data_rows = []
    for d, day_name in enumerate(day_order):
        for p, interval_str_val in enumerate(intervals_str):
            row = {
                "Day": day_name,
                "Interval": interval_str_val,
                "Inflated Required": inflated_req_matrix[d, p],
                "Scheduled": sched_matrix[d, p],
                "Over/(Under) vs Inflated": diff_matrix[d, p]
            }
            if has_base_req:
                row["Base Required"] = base_req_matrix[d, p]
            data_rows.append(row)
    interval_df = pd.DataFrame(data_rows)
    # Reorder columns to be more logical
    if has_base_req:
        interval_df = interval_df[["Day", "Interval", "Base Required", "Inflated Required", "Scheduled", "Over/(Under) vs Inflated"]]


    tab_charts, tab_heatmap, tab_vto_ot, tab_data = st.tabs(["📊 Charts", "🔥 Heatmap", "💸 VTO/OT Needs", "📋 Raw Data"])


    with tab_charts:
        st.markdown("##### Base, Inflated & Scheduled Staff")
        # --- Full week view restored ---
        fig = make_subplots(rows=7, cols=1, shared_xaxes=True, vertical_spacing=0.03, subplot_titles=day_order)
        for d in range(7):
            if has_base_req:
                fig.add_trace(go.Scatter(x=intervals_str, y=base_req_matrix[d, :], mode='lines', name='Base Required', line=dict(color='gray', dash='dot')), row=d+1, col=1)
            fig.add_trace(go.Scatter(x=intervals_str, y=inflated_req_matrix[d, :], mode='lines', name='Inflated Required', line=dict(color='blue', dash='dash')), row=d+1, col=1)
            fig.add_trace(go.Scatter(x=intervals_str, y=sched_matrix[d, :], mode='lines', name='Scheduled', line=dict(color='green')), row=d+1, col=1)
        fig.update_layout(height=1400, title_text="Daily Required vs. Scheduled Staff Levels", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
        # Hide duplicate legends
        for trace in fig.data[3:]:
            trace.showlegend = False
        st.plotly_chart(fig, use_container_width=True, key=f"{key_prefix}_daily_plot")

    with tab_heatmap:
        st.markdown("##### Over/Understaffing Heatmap (vs. Inflated Requirement)")
        fig_diff = go.Figure(data=go.Heatmap(
            z=diff_matrix.T,
            x=day_order,
            y=intervals_str,
            colorscale='RdBu',
            zmid=0,
            hovertemplate='Day: %{x}<br>Time: %{y}<br>Difference: %{z}<extra></extra>'))
        fig_diff.update_layout(title="Difference between Scheduled and Inflated Requirement")
        fig_diff.update_yaxes(autorange='reversed')
        st.plotly_chart(fig_diff, use_container_width=True, key=f"{key_prefix}_diff_heatmap")

    with tab_vto_ot:
        st.markdown("##### VTO (Overstaffing) and OT (Understaffing) Analysis")
        st.info("This shows the total hours of VTO opportunities and OT hours needed, calculated against the inflated requirement.")

        vto_df = interval_df[interval_df['Over/(Under) vs Inflated'] > 0][['Day', 'Interval', 'Over/(Under) vs Inflated']].rename(columns={'Over/(Under) vs Inflated': 'Surplus Staff'}).copy()
        ot_df = interval_df[interval_df['Over/(Under) vs Inflated'] < 0][['Day', 'Interval', 'Over/(Under) vs Inflated']].rename(columns={'Over/(Under) vs Inflated': 'Staff Shortfall'}).copy()
        ot_df['Staff Shortfall'] = ot_df['Staff Shortfall'].abs()

        if not vto_df.empty:
            vto_daily_summary = vto_df.groupby('Day')['Surplus Staff'].sum() * interval_duration_hours
            vto_daily_summary.name = "VTO Hours"
        else:
            vto_daily_summary = pd.Series(dtype=float, name="VTO Hours")

        if not ot_df.empty:
            ot_daily_summary = ot_df.groupby('Day')['Staff Shortfall'].sum() * interval_duration_hours
            ot_daily_summary.name = "OT Needed Hours"
        else:
            ot_daily_summary = pd.Series(dtype=float, name="OT Needed Hours")

        summary_vto_ot_df = pd.concat([vto_daily_summary, ot_daily_summary], axis=1).reindex(day_order).fillna(0)

        st.dataframe(summary_vto_ot_df.style.format("{:.1f} Hours"), use_container_width=True)
        download_dataframe_csv(summary_vto_ot_df, f"{key_prefix}_vto_ot_summary")

        vto_col, ot_col = st.columns(2)
        with vto_col:
            if st.checkbox("Show Detailed VTO Intervals", key=f"vto_detail_check_{key_prefix}"):
                if vto_df.empty:
                    st.info("No overstaffed intervals found.")
                else:
                    st.dataframe(vto_df, use_container_width=True, height=300)
                    download_dataframe_csv(vto_df, f"{key_prefix}_vto_details")

        with ot_col:
            if st.checkbox("Show Detailed OT Intervals", key=f"ot_detail_check_{key_prefix}"):
                if ot_df.empty:
                    st.info("No understaffed intervals found.")
                else:
                    st.dataframe(ot_df, use_container_width=True, height=300)
                    download_dataframe_csv(ot_df, f"{key_prefix}_ot_details")

    with tab_data:
        st.markdown("##### Interval-Level Data")
        st.dataframe(interval_df, use_container_width=True, height=500)
        download_dataframe_csv(interval_df, f"{key_prefix}_interval_data")

# ==============================================================================

# --- NEW: Comprehensive Save/Load Configuration using CSV ---
def get_app_config_for_csv():
    """Gathers all relevant state into a list of key-value pairs for CSV export."""
    config_rows = []

    # Helper to add a row to our list, serializing values appropriately
    def add_row(key, value):
        serialized_value = ""
        if value is None:
            serialized_value = "" # Represent None as an empty string
        elif isinstance(value, pd.DataFrame):
            # Serialize DataFrames to a JSON string
            serialized_value = value.to_json(orient='split', date_format='iso')
        elif isinstance(value, (list, dict)):
            # Serialize lists and dicts to a JSON string
            serialized_value = json.dumps(value, default=str)
        elif isinstance(value, (datetime.datetime, datetime.date, datetime.time, pd.Timestamp)):
            # Use ISO format for date/time objects
            serialized_value = value.isoformat()
        else:
            # For simple types (int, float, bool, str), just convert to string
            serialized_value = str(value)

        config_rows.append({'parameter': key, 'value': serialized_value})

    # Master list of all session state keys that represent user inputs to be saved.
    keys_to_save = [
        't1_start', 't1_end', 'week_start_day', 'working_hours', 'working_days',
        'base_hourly_rate', 'ot_hours_threshold', 'ot_rate_multiplier',
        'holiday_dates', 'holiday_prem_name', 'holiday_prem_mult',
        'sunday_pay_check', 'sunday_prem_mult', 'max_consecutive_slider',
        'min_off_days', 'understaff_penalty', 'overstaff_penalty', 'enable_adherence',
        'adherence_target_level', 'adherence_target_percent', 'adherence_cap_percent',
        'calc_mode', 'num_scenarios', 'num_blend_scen',
        't2_input_source', 'sched_mode', 'max_unique_shifts',
        'daily_op_hours',
        'allowed_durations', 'total_hc_optimization',
        'duration_rules', 'distribution_caps',
        'min_agents_per_shift', 'max_agents_per_shift',
        'optimization_model_choice', 'opt_adherence_target_level',
        'opt_adherence_target_percent', 'opt_adherence_cap_percent',
        'understaff_penalty_opt', 'overstaff_penalty_opt', 'shift_consistency_opt',
        'force_schedule_insufficient_hc' # NEW: Save this flag
    ]

    for key in keys_to_save:
        if key in st.session_state:
            add_row(key, st.session_state[key])

    # Handle dynamically generated widgets by saving the underlying data structures
    work_days_by_shift_data = {}
    if 'shifts_df' in st.session_state:
        for shift_name in st.session_state.shifts_df['Shift Name'].dropna().unique():
            key = f"work_days_{sanitize_name(shift_name)}"
            if key in st.session_state:
                work_days_by_shift_data[shift_name] = st.session_state[key]
    add_row('work_days_by_shift', work_days_by_shift_data)

    single_channel_scenarios_data = []
    num_scen = st.session_state.get('num_scenarios', 0)
    for i in range(num_scen):
        scen_data = {
            'name': st.session_state.get(f"scen_name_{i}"), 'type': st.session_state.get(f"scen_type_{i}"),
            'aht': st.session_state.get(f"scen_aht_{i}"), 'shrinkage': st.session_state.get(f"scen_shrink_{i}"),
            'target': st.session_state.get(f"scen_target_{i}") or st.session_state.get(f"scen_chat_target_{i}"),
            'awt': st.session_state.get(f"scen_awt_{i}") or st.session_state.get(f"scen_chat_awt_{i}"),
            'max_occupancy': st.session_state.get(f"scen_occ_{i}") or st.session_state.get(f"scen_chat_occ_{i}"),
            'concurrency': st.session_state.get(f"scen_concur_{i}"), 'volume_adjustment': st.session_state.get(f"scen_vol_adj_{i}"),
        }
        single_channel_scenarios_data.append(scen_data)
    add_row('single_channel_scenarios', single_channel_scenarios_data)

    # Explicitly save DataFrames
    for df_key in ['shifts_df', 'shift_differentials_df', 'single_channel_df', 'manual_req_df', 'daily_shrinkage_df']:
        if df_key in st.session_state and isinstance(st.session_state[df_key], pd.DataFrame):
            add_row(df_key, st.session_state[df_key])

    # Handle Blended Volumes by creating unique keys for each DataFrame
    if 'blended_volumes' in st.session_state:
        for (scen_idx, ch_name), df in st.session_state.blended_volumes.items():
            key = f"blended_volume__{scen_idx}__{ch_name}" # Use double underscore as a robust separator
            add_row(key, df)

    return pd.DataFrame(config_rows)

def load_app_config_from_csv(config_df):
    """Populates session_state from a loaded configuration DataFrame."""
    config_dict = config_df.set_index('parameter')['value'].to_dict()

    # Defines how to convert the string value from CSV back to its original type
    TYPE_CASTERS = {
        't1_start': lambda v: datetime.date.fromisoformat(v), 't1_end': lambda v: datetime.date.fromisoformat(v),
        'working_hours': float, 'working_days': float, 'base_hourly_rate': float, 'ot_hours_threshold': int,
        'ot_rate_multiplier': float, 'holiday_prem_mult': float, 'sunday_prem_mult': float,
        'max_consecutive_slider': int, 'min_off_days': int, 'understaff_penalty': int, 'overstaff_penalty': int,
        'adherence_target_percent': int, 'adherence_cap_percent': int,
        'sunday_pay_check': lambda v: v.lower() == 'true', 'enable_adherence': lambda v: v.lower() == 'true',
        'shift_consistency_opt': lambda v: v.lower() == 'true',
        'force_schedule_insufficient_hc': lambda v: v.lower() == 'true', # NEW: Load this flag
        'num_scenarios': int, 'num_blend_scen': int, 'max_unique_shifts': int, 'total_hc_optimization': int,
        'min_agents_per_shift': int, 'max_agents_per_shift': int, 'understaff_penalty_opt': int,
        'overstaff_penalty_opt': int, 'opt_adherence_target_percent': int, 'opt_adherence_cap_percent': int
    }

    # Keys for values that were stored as JSON strings (lists, dicts, DataFrames)
    JSON_LOAD_KEYS = [
        'holiday_dates', 'allowed_durations', 'duration_rules', 'distribution_caps',
        'work_days_by_shift', 'single_channel_scenarios', 'shifts_df', 'shift_differentials_df',
        'single_channel_df', 'manual_req_df', 'daily_shrinkage_df',
        'daily_op_hours'
    ]

    blended_volumes_to_load = {}
    for key, value_str in config_dict.items():
        if pd.isna(value_str) or value_str == '':
            st.session_state[key] = None
            continue

        if key.startswith('blended_volume__'):
            _, scen_idx_str, ch_name = key.split('__', 2)
            df = pd.read_json(StringIO(value_str), orient='split')
            blended_volumes_to_load[(int(scen_idx_str), ch_name)] = df
            continue

        if key in JSON_LOAD_KEYS:
            if 'df' in key:
                df = pd.read_json(StringIO(value_str), orient='split')

                # FIX: Correctly convert datetime columns to time objects without warnings
                if key in ['shifts_df', 'shift_differentials_df']:
                    if 'Start Time' in df.columns and pd.api.types.is_datetime64_any_dtype(df['Start Time']):
                        df['Start Time'] = df['Start Time'].dt.time
                    if 'End Time' in df.columns and pd.api.types.is_datetime64_any_dtype(df['End Time']):
                        df['End Time'] = df['End Time'].dt.time
                st.session_state[key] = df
            else:
                st.session_state[key] = json.loads(value_str)
        elif key in TYPE_CASTERS:
            st.session_state[key] = TYPE_CASTERS[key](value_str)
        else:
            st.session_state[key] = value_str

    if 'daily_op_hours' in st.session_state and st.session_state.daily_op_hours:
        loaded_dict = st.session_state.daily_op_hours
        converted_dict = {}
        for day, times in loaded_dict.items():
            converted_dict[day] = {
                'Start Time': datetime.time.fromisoformat(times.get('Start Time') or times.get('start')),
                'End Time': datetime.time.fromisoformat(times.get('End Time') or times.get('end'))
            }
        st.session_state.daily_op_hours = converted_dict

    st.session_state.blended_volumes = blended_volumes_to_load

    if 'work_days_by_shift' in st.session_state and st.session_state.work_days_by_shift:
        for shift_name, work_days in st.session_state.work_days_by_shift.items():
            st.session_state[f"work_days_{sanitize_name(shift_name)}"] = work_days

    if 'single_channel_scenarios' in st.session_state and st.session_state.single_channel_scenarios:
        scenarios_list = st.session_state.single_channel_scenarios
        st.session_state['num_scenarios'] = len(scenarios_list)
        for i, scen_data in enumerate(scenarios_list):
            st.session_state[f"scen_name_{i}"]=scen_data.get('name'); st.session_state[f"scen_type_{i}"]=scen_data.get('type')
            st.session_state[f"scen_aht_{i}"]=scen_data.get('aht'); st.session_state[f"scen_shrink_{i}"]=scen_data.get('shrinkage')
            st.session_state[f"scen_target_{i}"]=scen_data.get('target'); st.session_state[f"scen_chat_target_{i}"]=scen_data.get('target')
            st.session_state[f"scen_awt_{i}"]=scen_data.get('awt'); st.session_state[f"scen_chat_awt_{i}"]=scen_data.get('awt')
            st.session_state[f"scen_occ_{i}"]=scen_data.get('max_occupancy'); st.session_state[f"scen_chat_occ_{i}"]=scen_data.get('max_occupancy')
            st.session_state[f"scen_concur_{i}"]=scen_data.get('concurrency'); st.session_state[f"scen_vol_adj_{i}"]=scen_data.get('volume_adjustment')


# ----------------- UI Starts Here -----------------

start_idx = DAYS_OF_WEEK_OPTIONS.index(st.session_state.get('week_start_day', "Sunday"))
days_of_week_ordered = DAYS_OF_WEEK_OPTIONS[start_idx:] + DAYS_OF_WEEK_OPTIONS[:start_idx]

# Apply safe update to daily shrinkage dataframe
intervals_str_index = [t.strftime('%H:%M') for t in st.session_state.intervals]
if 'daily_shrinkage_df' not in st.session_state or list(st.session_state.daily_shrinkage_df.columns) != days_of_week_ordered:
    safe_update_dataframe('daily_shrinkage_df', days_of_week_ordered, intervals_str_index)
st.session_state.daily_shrinkage_df.index.name = "Interval"


with st.sidebar.expander("📲 Configuration Management", expanded=True):
    st.info("Save all settings from the sidebar and Tab 1 & 2 to a single CSV file, or load a previous configuration.", icon="ℹ️")

    # Save/Export configuration to CSV
    try:
        config_df = get_app_config_for_csv()
        csv_data = config_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Download Configuration",
            data=csv_data,
            file_name=f"wfm_config_{datetime.date.today()}.csv",
            mime="text/csv",
            help="Downloads all current settings into a single CSV file."
        )
    except Exception as e:
        st.error(f"Could not generate configuration CSV: {e}")

    # Load/Import configuration from CSV
    uploaded_config_csv = st.file_uploader(
        "Upload Configuration File",
        type="csv",
        key="config_uploader_csv",
        help="Upload a CSV file previously downloaded from this app to restore all settings."
    )
    if uploaded_config_csv is not None:
        try:
            loaded_df = pd.read_csv(uploaded_config_csv)
            load_app_config_from_csv(loaded_df)
            st.success("Configuration loaded successfully! The app will now update with the new settings.")
            st.rerun()
        except Exception as e:
            st.error(f"Error loading configuration from CSV: {e}")
            st.exception(e)


tab1, tab2, tab3, tab4 = st.tabs(["1. Demand & Staffing Forecast", "2. Schedule & Cost Simulation", "3. Results Summary", "4. Forecast 'What-If' Simulation"])

with tab1:
    st.header("Step 1: Calculate Staffing Requirements")
    st.info("Forecast the number of staff required per interval based on your workload and service targets. You can calculate for a single channel or create a blended forecast for multiple channels.")

    with st.sidebar.expander("🗓️ Date & Week Configuration", expanded=True):
        today = datetime.date.today()
        # Ensure widgets use session state for load functionality
        start_date = st.date_input("Start Date", value=st.session_state.get('t1_start', today), key="t1_start")
        end_date = st.date_input("End Date", value=st.session_state.get('t1_end', today + datetime.timedelta(days=6)), key="t1_end")
        
        # Robust index calculation for selectbox
        saved_week_start_day = st.session_state.get('week_start_day', 'Sunday')
        try:
            week_start_index = DAYS_OF_WEEK_OPTIONS.index(saved_week_start_day)
        except ValueError:
            week_start_index = 0 # Default to first option if saved value is invalid
        week_start_day_name_input = st.selectbox("Week Starts On", DAYS_OF_WEEK_OPTIONS, index=week_start_index, key="week_start_day")

        if start_date > end_date: st.error("Error: End date must fall after start date."); st.stop()

    with st.sidebar.expander("⚙️ General WFM Parameters", expanded=False):
        st.number_input("Typical Working Hours per Day", min_value=1.0, max_value=24.0, value=st.session_state.get('working_hours', 8.0), step=0.5, key="working_hours")
        st.number_input("Typical Working Days per Week", min_value=1.0, max_value=7.0, value=st.session_state.get('working_days', 5.0), step=0.5, key="working_days")


    input_date_range = pd.date_range(start_date, end_date)
    input_dates_str = input_date_range.strftime('%Y-%m-%d').tolist()
    day_name_map = {date.strftime('%Y-%m-%d'): date.strftime('%A') for date in input_date_range}
    intervals_list = st.session_state.intervals
    interval_index_str = [t.strftime('%H:%M:%S') for t in intervals_list]

    st.markdown("---")
    st.markdown("#### **Select Calculation Mode**")
    
    calc_mode_options = ("Single Channel (Run Multiple Scenarios)", "Blended (Multi-Channel)")
    saved_calc_mode = st.session_state.get('calc_mode', calc_mode_options[0])
    try:
        calc_mode_index = calc_mode_options.index(saved_calc_mode)
    except ValueError:
        calc_mode_index = 0
    calc_mode = st.radio(
        "How do you want to calculate staffing?",
        calc_mode_options,
        index=calc_mode_index,
        key="calc_mode",
        horizontal=True,
        label_visibility="collapsed"
    )

    if calc_mode == "Single Channel (Run Multiple Scenarios)":
        st.markdown("#### Define Scenarios")
        st.info("Set the number of 'what-if' scenarios to run. Each scenario can have different parameters and its own volume adjustment, but all use the same base volume data entered at the bottom.")

        num_scenarios = st.number_input("How many scenarios do you want to model?", min_value=1, value=st.session_state.get('num_scenarios', 2), step=1, key="num_scenarios")

        scenarios_to_run = []
        for i in range(num_scenarios):
            st.markdown(f"---")
            st.markdown(f"##### Parameters for Scenario #{i+1}")

            cols1 = st.columns([2, 2, 1, 1])
            scenario_name = cols1[0].text_input("Scenario Name", value=st.session_state.get(f'scen_name_{i}', f"Scenario {i+1}"), key=f"scen_name_{i}")

            channel_options = ["Voice (Erlang-C)", "Chat (Erlang with Concurrency)", "Email / Back Office (Transactional)"]
            default_channel = st.session_state.get(f'scen_type_{i}', "Voice (Erlang-C)")
            try:
                default_channel_idx = channel_options.index(default_channel)
            except ValueError:
                default_channel_idx = 0

            channel_type = cols1[1].selectbox(
                "Channel Type",
                options=channel_options,
                key=f"scen_type_{i}",
                index=default_channel_idx
            )
            aht = cols1[2].number_input("AHT (s)", min_value=1, value=st.session_state.get(f'scen_aht_{i}', 300), key=f"scen_aht_{i}")
            shrinkage = cols1[3].number_input("Shrinkage (%)", min_value=0.0, max_value=99.9, value=st.session_state.get(f'scen_shrink_{i}', 30.0), step=0.5, key=f"scen_shrink_{i}")

            params = {
                'scenario_name': scenario_name,
                'channel_type': channel_type,
                'aht': aht,
                'shrinkage': shrinkage
            }

            if channel_type == "Voice (Erlang-C)":
                cols2 = st.columns(3)
                params['target'] = cols2[0].number_input("SL Target (%)", min_value=1.0, max_value=100.0, value=st.session_state.get(f'scen_target_{i}', 80.0), step=0.5, key=f"scen_target_{i}")
                params['awt'] = cols2[1].number_input("AWT (s)", min_value=1, value=st.session_state.get(f'scen_awt_{i}', 20), key=f"scen_awt_{i}")
                params['max_occupancy'] = cols2[2].number_input("Max Occupancy (%)", min_value=1.0, max_value=100.0, value=st.session_state.get(f'scen_occ_{i}', 85.0), step=0.1, key=f"scen_occ_{i}")

            elif channel_type == "Chat (Erlang with Concurrency)":
                st.info("This model uses Erlang C, adjusting for agent concurrency. It is ideal for chat channels with specific service level goals.")
                cols2 = st.columns(4)
                params['target'] = cols2[0].number_input("SL Target (%)", min_value=1.0, max_value=100.0, value=st.session_state.get(f'scen_chat_target_{i}', 80.0), step=0.5, key=f"scen_chat_target_{i}")
                params['awt'] = cols2[1].number_input("AWT (s)", min_value=1, value=st.session_state.get(f'scen_chat_awt_{i}', 30), key=f"scen_chat_awt_{i}")
                params['max_occupancy'] = cols2[2].number_input("Max Occupancy (%)", min_value=1.0, max_value=100.0, value=st.session_state.get(f'scen_chat_occ_{i}', 85.0), step=0.1, key=f"scen_chat_occ_{i}")
                params['concurrency'] = cols2[3].number_input("Agent Concurrency", min_value=1.0, value=st.session_state.get(f'scen_concur_{i}', 3.0), step=0.1, key=f"scen_concur_{i}", help="How many chats an agent can handle at the same time.")

            params['volume_adjustment'] = st.number_input(
                "Volume as % of Input",
                min_value=0.0,
                max_value=500.0,
                value=st.session_state.get(f'scen_vol_adj_{i}', 100.0),
                step=5.0,
                key=f"scen_vol_adj_{i}",
                help="Set this scenario's volume as a percentage of the base input. 100% = no change, 0% = zero volume, 200% = double volume."
            )
            scenarios_to_run.append(params)

        st.markdown("---")
        st.markdown("#### Base Workload Volume (Used for all scenarios above)")
        
        # Apply safe update to single channel dataframe
        if "single_channel_df" not in st.session_state or list(st.session_state["single_channel_df"].columns) != input_dates_str:
            safe_update_dataframe("single_channel_df", input_dates_str, interval_index_str)

        st.session_state["single_channel_df"] = st.data_editor(st.session_state["single_channel_df"], key="single_channel_editor", height=300, use_container_width=True)
        download_dataframe_csv(st.session_state["single_channel_df"], "base_workload_volume")

        if st.button("Calculate Staffing for All Scenarios", type="primary", key="calc_all_scenarios"):
            start_time = time.time()
            base_volume_df = st.session_state.single_channel_df.copy()
            all_scenarios_results = {}
            all_summary_rows = []
            has_errors = False

            scenario_names = [s['scenario_name'] for s in scenarios_to_run]
            if len(scenario_names) != len(set(scenario_names)):
                st.error("Scenario names must be unique. Please assign a different name to each scenario.")
                st.stop()

            with st.spinner(f"Calculating {len(scenarios_to_run)} scenarios..."):
                progress_bar = st.progress(0)
                for i, params in enumerate(scenarios_to_run):
                    scenario_name = params['scenario_name']
                    try:
                        vol_adj_percent = params.get('volume_adjustment', 100.0)
                        # Ensure base_volume_df is numeric before multiplication
                        numeric_base_volume_df = base_volume_df.apply(pd.to_numeric, errors='coerce').fillna(0)
                        adjusted_volume_df = numeric_base_volume_df * (vol_adj_percent / 100.0)
                        # Store adjusted volume for potential use in Tab 4
                        params['base_volume_df'] = numeric_base_volume_df

                        staffing_df = run_staffing_calculation(params, input_dates_str, day_name_map, st.session_state.week_start_day, adjusted_volume_df)
                        all_scenarios_results[scenario_name] = (staffing_df, params)

                        for week_start_date in staffing_df['Week_Start_Day'].unique():
                            weekly_df = staffing_df[staffing_df['Week_Start_Day'] == week_start_date].copy()
                            weekly_kpis = calculate_aggregated_kpis(weekly_df)
                            req_pivot_table = weekly_df.pivot_table(index='Date', columns='Interval', values='final_positions', fill_value=0)
                            start_date_obj = week_start_date.date() if hasattr(week_start_date, 'date') else week_start_date
                            week_dates_for_pivot = pd.to_datetime([(start_date_obj + datetime.timedelta(days=d)) for d in range(7)])
                            req_pivot_table = req_pivot_table.reindex(index=week_dates_for_pivot, fill_value=0)
                            fte_metrics = calculate_fte_metrics_from_matrix(
                                req_pivot_table.values.tolist(), st.session_state.working_hours, st.session_state.working_days
                            )

                            all_summary_rows.append({
                                "Scenario": scenario_name, "Week_Start_Day": pd.to_datetime(week_start_date).strftime('%Y-%m-%d'),
                                "Required HC (Avg)": fte_metrics['avg_fte'], "Required HC for Peak Day": fte_metrics['peak_fte'],
                                "Total Volume": weekly_kpis["Total Calls"], "Weekly ASA (s)": weekly_kpis["Overall ASA (s)"],
                                "Weekly SL (%)": weekly_kpis["Service Level (%)"], "Weekly Occ. (%)": weekly_kpis["Occupancy (%)"],
                                "Weekly Wait Prob. (%)": weekly_kpis["Wait Probability (%)"]
                            })
                    except Exception as e:
                        st.error(f"Error processing scenario '{scenario_name}': {e}")
                        has_errors = True
                    progress_bar.progress((i + 1) / len(scenarios_to_run))

            st.session_state['all_scenarios'] = all_scenarios_results
            st.session_state['scenario_summary'] = pd.DataFrame(all_summary_rows) if all_summary_rows else pd.DataFrame()

            end_time = time.time()
            duration = end_time - start_time
            formatted_time = format_duration(duration)

            if not has_errors: st.success(f"All scenarios calculated successfully! Time taken: {formatted_time}")
            else: st.warning(f"Some scenarios failed to calculate. Time taken: {formatted_time}")

    else: # Blended Mode
        st.markdown("#### Blended Channel Scenarios")
        st.info("Define multiple blended scenarios. Each scenario can have a different name, mix of channels, parameters, and workloads.")
        num_blend_scenarios = st.number_input("How many blended scenarios to model?", 1, 10, value=st.session_state.get('num_blend_scen', 1), key="num_blend_scen")

        CHANNEL_OPTIONS = {
            "Voice": "Voice (Erlang-C)", "Chat": "Chat (Erlang with Concurrency)", "Email/BO": "Email / Back Office (Transactional)"
        }
        all_blend_scenarios_params = []
        
        # Apply safe update to blended volumes dictionary
        for vol_key in list(st.session_state.blended_volumes.keys()):
            df = st.session_state.blended_volumes[vol_key]
            if list(df.columns) != input_dates_str:
                temp_key = f"blended_volume__{vol_key[0]}__{vol_key[1]}"
                safe_update_dataframe(temp_key, input_dates_str, interval_index_str)
                st.session_state.blended_volumes[vol_key] = st.session_state.pop(temp_key)


        for i in range(num_blend_scenarios):
            with st.container(border=True):
                st.markdown(f"### Blended Scenario #{i+1}")
                scenario_params = {}

                cols = st.columns([2,1])
                scenario_params['name'] = cols[0].text_input("Scenario Name", f"Blended Scenario {i+1}", key=f"b_scen_name_{i}")
                scenario_params['vol_adjust'] = cols[1].number_input(
                    "Volume as % of Input",
                    min_value=0.0,
                    max_value=500.0,
                    value=100.0,
                    step=5.0,
                    key=f"b_vol_adjust_{i}",
                    help="Set this scenario's volume as a percentage of the base input. 100% = no change, 0% = zero volume, 200% = double volume."
                )

                scenario_params['channels'] = st.multiselect(
                    "Select channels to blend for this scenario:", options=list(CHANNEL_OPTIONS.keys()), default=list(CHANNEL_OPTIONS.keys()), key=f"b_channels_{i}"
                )
                is_erlang_only_blend = scenario_params['channels'] and "Email/BO" not in scenario_params['channels']
                scenario_params['erlang_only'] = is_erlang_only_blend

                scenario_params['channel_params'] = {}
                scenario_params['volume_dfs'] = {}

                if scenario_params['channels']:
                    if is_erlang_only_blend:
                        with st.container(border=True):
                            st.markdown("##### 🎯 Blended KPI Targets")
                            b_cols = st.columns(3)
                            scenario_params['target'] = b_cols[0].number_input("Blended SL Target (%)", 1.0, 100.0, 80.0, 0.5, key=f"b_sl_{i}")
                            scenario_params['awt'] = b_cols[1].number_input("Blended AWT (s)", 1, 300, 25, key=f"b_awt_{i}")
                            scenario_params['max_occupancy'] = b_cols[2].number_input("Blended Max Occupancy (%)", 1.0, 100.0, 85.0, 0.1, key=f"b_occ_{i}")

                    tabs = st.tabs(scenario_params['channels'])
                    for j, short_name in enumerate(scenario_params['channels']):
                        with tabs[j]:
                            st.markdown(f"###### Parameters & Workload for: **{short_name}**")
                            p = {}
                            if short_name == "Voice":
                                v_cols = st.columns([1,1])
                                p['aht'] = v_cols[0].number_input("AHT (s)", 1, value=300, key=f"v_aht_{i}_{j}")
                                p['shrinkage'] = v_cols[1].number_input("Shrinkage (%)", 0.0, 99.0, 30.0, 0.5, key=f"v_sh_{i}_{j}")
                            elif short_name == "Chat":
                                c_cols = st.columns(3)
                                p['aht'] = c_cols[0].number_input("AHT (s)", 1, value=600, key=f"c_aht_{i}_{j}")
                                p['shrinkage'] = c_cols[1].number_input("Shrinkage (%)", 0.0, 99.0, 25.0, 0.5, key=f"c_sh_{i}_{j}")
                                p['concurrency'] = c_cols[2].number_input("Concurrency", 1.0, 10.0, 2.5, 0.1, key=f"c_con_{i}_{j}")
                            elif short_name == "Email/BO":
                                e_cols = st.columns(2)
                                p['aht'] = e_cols[0].number_input("AHT (s)", 1, value=450, key=f"e_aht_{i}_{j}")
                                p['shrinkage'] = e_cols[1].number_input("Shrinkage (%)", 0.0, 99.0, 20.0, 0.5, key=f"e_sh_{i}_{j}")

                            scenario_params['channel_params'][short_name] = p
                            vol_key = (i, short_name)
                            if vol_key not in st.session_state.blended_volumes:
                                st.session_state.blended_volumes[vol_key] = pd.DataFrame(0.0, index=interval_index_str, columns=input_dates_str)
                            
                            st.session_state.blended_volumes[vol_key] = st.data_editor(st.session_state.blended_volumes[vol_key], key=f"vol_editor_{i}_{j}", height=300)
                            download_dataframe_csv(st.session_state.blended_volumes[vol_key], f"blended_vol_{scenario_params['name']}_{short_name}")
                            scenario_params['volume_dfs'][short_name] = st.session_state.blended_volumes[vol_key]
                all_blend_scenarios_params.append(scenario_params)

        if st.button("Calculate All Blended Scenarios", type="primary", key="calc_blended_scenarios"):
            start_time = time.time()
            all_scenarios_results = {}
            all_summary_rows = []
            has_errors = False

            scenario_names = [s['name'] for s in all_blend_scenarios_params]
            if len(scenario_names) != len(set(scenario_names)):
                st.error("Blended scenario names must be unique.")
                st.stop()

            interval_seconds = (pd.to_timedelta(st.session_state.interval_freq).total_seconds())

            with st.spinner("Calculating all blended scenarios..."):
                for scen_params in all_blend_scenarios_params:
                    scenario_name = scen_params['name']
                    try:
                        vol_adjust_percent = scen_params.get('vol_adjust', 100.0)
                        adjusted_volume_dfs = {}
                        for ch, df in scen_params['volume_dfs'].items():
                            adjusted_volume_dfs[ch] = df.copy() * (vol_adjust_percent / 100.0)

                        staffing_df = pd.DataFrame()
                        if scen_params.get('erlang_only', False):
                            total_workload_df = pd.DataFrame(0.0, index=interval_index_str, columns=input_dates_str)
                            total_volume_df = pd.DataFrame(0.0, index=interval_index_str, columns=input_dates_str)
                            total_vol_x_shrinkage = 0
                            for ch_name in scen_params['channels']:
                                params = scen_params['channel_params'][ch_name]
                                vol_df = adjusted_volume_dfs[ch_name]
                                total_volume_df += vol_df
                                workload = vol_df.multiply(params['aht'] / params.get('concurrency', 1.0))
                                total_workload_df += workload
                                total_vol_x_shrinkage += (vol_df * params['shrinkage']).sum().sum()
                            blended_aht_df = total_workload_df.divide(total_volume_df).fillna(0)
                            blended_shrinkage = (total_vol_x_shrinkage / total_volume_df.sum().sum()) if total_volume_df.sum().sum() > 0 else 0
                            scen_params['shrinkage'] = blended_shrinkage

                            blend_staffing_results = []
                            for date_str in input_dates_str:
                                for i_idx, interval_str in enumerate(interval_index_str):
                                    interval_time_obj = intervals_list[i_idx]
                                    volume = total_volume_df.loc[interval_str, date_str]
                                    aht = blended_aht_df.loc[interval_str, date_str]
                                    temp_params = {'channel_type': 'Voice (Erlang-C)', 'aht': aht, **scen_params}
                                    common_data = {"Date": pd.to_datetime(date_str), "Day": day_name_map[date_str], "Interval": interval_time_obj, "Week_Start_Day": get_week_start(datetime.datetime.strptime(date_str, '%Y-%m-%d'), st.session_state.week_start_day), "Volume": volume, "AHT": aht}

                                    if volume > 0:
                                        kpi_results = calculate_erlang_c_positions(temp_params['awt'], temp_params['shrinkage'], temp_params['max_occupancy'], aht, temp_params['target'], volume)
                                        kpis = kpi_results[0]
                                        raw_positions_needed = kpis['positions']

                                        awt_for_queued = kpis.get('asa', 0)
                                        wp = kpis.get('waiting_probability', 0)
                                        result_row = {
                                            'raw_positions': raw_positions_needed,
                                            'final_positions': math.ceil(raw_positions_needed * (1 / (1- (temp_params['shrinkage']/100)))),
                                            'service_level': kpis.get('service_level', 0),
                                            'occupancy': kpis.get('occupancy', 0),
                                            'waiting_probability': wp,
                                            'AWT_for_Queued_s': awt_for_queued,
                                            'ASA_s': awt_for_queued * wp
                                        }
                                        blend_staffing_results.append({**result_row, **common_data})
                                    else:
                                        blend_staffing_results.append({**common_data, "raw_positions": 0, "final_positions": 0, "service_level": 1.0, "occupancy": 0.0, "waiting_probability": 0.0, "AWT_for_Queued_s": 0.0, "ASA_s": 0.0})
                            staffing_df = pd.DataFrame(blend_staffing_results)
                        else: # Sum of parts
                            total_reqs_df = pd.DataFrame(0.0, index=interval_index_str, columns=input_dates_str)
                            total_volume_df = pd.DataFrame(0.0, index=interval_index_str, columns=input_dates_str)
                            for ch_name in scen_params['channels']:
                                params = scen_params['channel_params'][ch_name]
                                params['channel_type'] = CHANNEL_OPTIONS[ch_name]
                                vol_df = adjusted_volume_dfs[ch_name]
                                total_volume_df += vol_df
                                channel_staffing_df = run_staffing_calculation(params, input_dates_str, day_name_map, st.session_state.week_start_day, vol_df)
                                req_pivot = channel_staffing_df.pivot_table(index='Interval', columns='Date', values='final_positions', fill_value=0)
                                req_pivot.index = [t.strftime('%H:%M:%S') for t in req_pivot.index]
                                req_pivot.columns = req_pivot.columns.strftime('%Y-%m-%d')
                                req_pivot = req_pivot.reindex(index=interval_index_str, columns=input_dates_str).fillna(0)
                                total_reqs_df += req_pivot
                            total_reqs_long = total_reqs_df.reset_index().melt(id_vars='index', var_name='Date', value_name='final_positions')
                            total_reqs_long.rename(columns={'index': 'Interval_str'}, inplace=True)
                            total_reqs_long['Date'] = pd.to_datetime(total_reqs_long['Date'])
                            time_map = {t.strftime('%H:%M:%S'): t for t in intervals_list}
                            total_reqs_long['Interval'] = total_reqs_long['Interval_str'].map(time_map)
                            total_volume_long = total_volume_df.reset_index().melt(id_vars='index', var_name='Date', value_name='Volume')
                            total_reqs_long['Volume'] = total_volume_long['Volume']
                            staffing_df = total_reqs_long[['Date', 'Interval', 'Week_Start_Day', 'final_positions', 'Volume']]

                        all_scenarios_results[scenario_name] = (staffing_df, scen_params)
                        for week_start_date in staffing_df['Week_Start_Day'].unique():
                            weekly_df = staffing_df[staffing_df['Week_Start_Day'] == week_start_date].copy()
                            req_pivot_table = weekly_df.pivot_table(index='Date', columns='Interval', values='final_positions', fill_value=0)
                            start_date_obj = week_start_date.date()
                            week_dates = pd.to_datetime([(start_date_obj + datetime.timedelta(days=d)) for d in range(7)])
                            req_pivot_table = req_pivot_table.reindex(index=week_dates, fill_value=0)
                            fte_metrics = calculate_fte_metrics_from_matrix(req_pivot_table.values.tolist(), st.session_state.working_hours, st.session_state.working_days)
                            summary = {"Scenario": scenario_name, "Week_Start_Day": week_start_date.strftime('%Y-%m-%d'), "Required HC (Avg)": fte_metrics['avg_fte'], "Required HC for Peak Day": fte_metrics['peak_fte'], "Total Volume": int(weekly_df['Volume'].sum())}
                            if 'service_level' in weekly_df.columns:
                                weekly_kpis = calculate_aggregated_kpis(weekly_df)
                                summary.update({"Weekly ASA (s)": weekly_kpis["Overall ASA (s)"], "Weekly SL (%)": weekly_kpis["Service Level (%)"], "Weekly Occ. (%)": weekly_kpis["Occupancy (%)"], "Weekly Wait Prob. (%)": weekly_kpis["Wait Probability (%)"]})
                            all_summary_rows.append(summary)
                    except Exception as e:
                        st.error(f"Error in blended scenario '{scenario_name}': {e}")
                        has_errors = True

            st.session_state['all_scenarios'] = all_scenarios_results
            st.session_state['scenario_summary'] = pd.DataFrame(all_summary_rows) if all_summary_rows else pd.DataFrame()

            end_time = time.time()
            duration = end_time - start_time
            formatted_time = format_duration(duration)

            if not has_errors: st.success(f"All blended scenarios calculated! Time taken: {formatted_time}")
            else: st.warning(f"Some blended scenarios failed. Time taken: {formatted_time}")

    if "scenario_summary" in st.session_state and not st.session_state.scenario_summary.empty:
        st.markdown("---")
        st.subheader("Scenario Summary Table (per Week)")
        st.dataframe(st.session_state.scenario_summary.style.format({
            "Required HC (Avg)": '{:.2f}',
            "Required HC for Peak Day": '{:.2f}',
            "Total Volume": '{:,.0f}',
            "Weekly SL (%)": '{:.2f}%',
            "Weekly Occ. (%)": '{:.2f}%',
        }, na_rep="N/A"), use_container_width=True)
        download_dataframe_csv(st.session_state.scenario_summary, "scenario_summary_table")


        if len(st.session_state.get('all_scenarios', {})) >= 2:
            st.markdown("---")
            st.subheader("Detailed Scenario Comparison")
            st.info("Select two scenarios and a common week to compare their required staffing levels and performance side-by-side, including the difference.")

            scenario_options = list(st.session_state.get('all_scenarios', {}).keys())
            summary_df = st.session_state.scenario_summary

            def get_scenario_data_for_week(scenario_name, week_str):
                week_summary_row = summary_df[(summary_df['Scenario'] == scenario_name) & (summary_df['Week_Start_Day'] == week_str)]
                if week_summary_row.empty:
                    return None

                staffing_df, _ = st.session_state['all_scenarios'][scenario_name]
                week_start_dt = pd.to_datetime(week_str)
                weekly_df = staffing_df[staffing_df['Week_Start_Day'] == week_start_dt].copy()
                if weekly_df.empty:
                    return None

                weekly_df['Day'] = pd.Categorical(weekly_df['Date'].dt.strftime('%A'), categories=days_of_week_ordered, ordered=True)

                req_pivot = weekly_df.pivot_table(index='Interval', columns='Day', values='final_positions', aggfunc='sum', fill_value=0)
                req_pivot = req_pivot.reindex(index=st.session_state.intervals, fill_value=0)
                req_pivot = req_pivot.reindex(columns=days_of_week_ordered, fill_value=0)
                daily_totals = req_pivot.sum()

                return {"summary": week_summary_row, "pivot": req_pivot, "daily_totals": daily_totals}

            def display_scenario_comparison_card(column, scenario_name, scenario_data, week_str, key_suffix):
                with column:
                    st.markdown(f"#### {scenario_name}")
                    st.caption(f"Week of: {week_str}")
                    card_df = scenario_data['summary'].drop(columns=['Scenario', 'Week_Start_Day']).T
                    card_df.columns = ['Value']
                    card_df['Value'] = card_df['Value'].apply(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)
                    st.dataframe(card_df, use_container_width=True)
                    download_dataframe_csv(card_df, f"compare_{key_suffix}_summary_card")

                    st.markdown("##### Daily Staffing Requirements by Interval")
                    st.dataframe(scenario_data['pivot'].style.background_gradient(cmap='viridis'), use_container_width=True)
                    download_dataframe_csv(scenario_data['pivot'], f"compare_{key_suffix}_staffing_requirements")

                    st.markdown("##### Total Required Staff-Intervals per Day")
                    st.dataframe(scenario_data['daily_totals'].apply(lambda x: f"{x:,.0f}").to_frame(name="Staff-Intervals"), use_container_width=True)
                    download_dataframe_csv(scenario_data['daily_totals'].apply(lambda x: f"{x:,.0f}").to_frame(name="Staff-Intervals"), f"compare_{key_suffix}_daily_totals")

            def display_scenario_difference_card(column, data_A, data_B, week_str, key_suffix, name_A, name_B):
                 with column:
                    title_name = f"Difference ({name_B} - {name_A})"
                    st.markdown(f"#### {title_name}")
                    st.caption(f"Week of: {week_str}")

                    summary_A = data_A['summary'].drop(columns=['Scenario', 'Week_Start_Day']).iloc[0].apply(pd.to_numeric, errors='coerce')
                    summary_B = data_B['summary'].drop(columns=['Scenario', 'Week_Start_Day']).iloc[0].apply(pd.to_numeric, errors='coerce')
                    summary_diff = summary_B - summary_A
                    summary_diff_df = summary_diff.to_frame(name=f"Difference ({name_B}-{name_A})")
                    st.dataframe(summary_diff_df.style.format("{:,.2f}", na_rep="-"), use_container_width=True)
                    download_dataframe_csv(summary_diff_df, f"compare_{key_suffix}_summary_diff")

                    pivot_diff = data_B['pivot'] - data_A['pivot']
                    st.markdown("##### Requirement Difference (Daily Interval Staffing)")
                    st.dataframe(pivot_diff.style.background_gradient(cmap='RdBu', axis=None), use_container_width=True)
                    download_dataframe_csv(pivot_diff, f"compare_{key_suffix}_staffing_difference")

                    daily_totals_diff = data_B['daily_totals'] - data_A['daily_totals']
                    st.markdown("##### Daily Staff-Intervals Difference")
                    st.dataframe(daily_totals_diff.apply(lambda x: f"{x:,.0f}").to_frame(name="Staff-Intervals Diff"), use_container_width=True)
                    download_dataframe_csv(daily_totals_diff.apply(lambda x: f"{x:,.0f}").to_frame(name="Staff-Intervals Diff"), f"compare_{key_suffix}_daily_totals_diff")


            sel_col1, sel_col2 = st.columns(2)
            selection_A = sel_col1.selectbox("Compare Scenario A", scenario_options, index=0, key="compare_A_scen")
            selection_B = sel_col2.selectbox("Compare Scenario B", scenario_options, index=1 if len(scenario_options) > 1 else 0, key="compare_B_scen")

            common_weeks = []
            if selection_A and selection_B and selection_A != selection_B:
                weeks_A = set(summary_df[summary_df['Scenario'] == selection_A]['Week_Start_Day'].unique())
                weeks_B = set(summary_df[summary_df['Scenario'] == selection_B]['Week_Start_Day'].unique())
                common_weeks = sorted(list(weeks_A.intersection(weeks_B)))

            selected_week_for_comparison = None
            if common_weeks:
                selected_week_for_comparison = st.selectbox("Select a common week to compare", options=common_weeks, key="compare_week_select")

            if selection_A and selection_B and selection_A != selection_B:
                if selected_week_for_comparison:
                    data_A = get_scenario_data_for_week(selection_A, selected_week_for_comparison)
                    data_B = get_scenario_data_for_week(selection_B, selected_week_for_comparison)

                    if data_A and data_B:
                        disp_col1, disp_col2, disp_col3 = st.columns(3)
                        display_scenario_comparison_card(disp_col1, selection_A, data_A, selected_week_for_comparison, "A")
                        display_scenario_comparison_card(disp_col2, selection_B, data_B, selected_week_for_comparison, "B")
                        display_scenario_difference_card(disp_col3, data_A, data_B, selected_week_for_comparison, "Diff", selection_A, selection_B)
                    else:
                        st.error("Could not load complete data for one or both scenarios for the selected week.")

                elif not common_weeks:
                    st.warning("These scenarios have no overlapping weeks to compare. Please calculate scenarios over a similar date range.")
            elif selection_A == selection_B:
                 st.warning("Please select two different scenarios to compare.")


        st.markdown("---")
        st.header("Interval Level Forecast Details")
        st.info("Select a calculated scenario and week to view the detailed interval-level forecast.")

        scenario_options_detail = list(st.session_state.get('all_scenarios', {}).keys())
        if not scenario_options_detail:
            st.warning("No scenarios have been calculated yet.")
        else:
            selected_scenario_for_detail = st.selectbox("Select a Scenario", options=scenario_options_detail, key="t1_detail_scenario")

            if selected_scenario_for_detail:
                scenario_data = st.session_state.all_scenarios[selected_scenario_for_detail]
                staffing_df, extra_data = scenario_data
                is_blended = isinstance(extra_data, dict) and 'channels' in extra_data

                staffing_df['Week_Start_Day'] = pd.to_datetime(staffing_df['Week_Start_Day'])
                week_start_options = sorted(staffing_df['Week_Start_Day'].dt.strftime('%Y-%m-%d').unique())

                if not week_start_options:
                    st.warning("No data available for the selected scenario in the chosen date range.")
                    st.stop()

                selected_week_for_detail = st.selectbox("Select a Week", options=week_start_options, key="t1_detail_week")

                if selected_week_for_detail:
                    week_start_dt = datetime.datetime.strptime(selected_week_for_detail, '%Y-%m-%d').date()

                    week_dates_str = [(week_start_dt + datetime.timedelta(days=d)).strftime('%Y-%m-%d') for d in range(7)]

                    weekly_df = staffing_df[staffing_df['Date'].dt.strftime('%Y-%m-%d').isin(week_dates_str)].copy()
                    weekly_df['Day'] = pd.Categorical(weekly_df['Date'].dt.strftime('%A'), categories=days_of_week_ordered, ordered=True)

                    req_pivot = weekly_df.pivot_table(index='Interval', columns='Day', values='final_positions', aggfunc='sum', fill_value=0)
                    volume_pivot = weekly_df.pivot_table(index='Interval', columns='Day', values='Volume', aggfunc='sum', fill_value=0)

                    req_pivot = req_pivot.reindex(index=st.session_state.intervals, fill_value=0)
                    volume_pivot = volume_pivot.reindex(index=st.session_state.intervals, fill_value=0)

                    req_pivot = req_pivot.reindex(columns=days_of_week_ordered, fill_value=0)
                    volume_pivot = volume_pivot.reindex(columns=days_of_week_ordered, fill_value=0)

                    intervals_str_fmt = [t.strftime('%H:%M') for t in st.session_state.intervals]

                    tab_charts, tab_data_details = st.tabs(["📊 Charts & Heatmaps", "📋 Detailed Data"])

                    with tab_charts:
                        st.markdown("##### Required Staff vs. Workload Volume")
                        if is_blended:
                             st.info("Showing blended (total) requirement.")
                        fig = make_subplots(rows=7, cols=1, shared_xaxes=True, vertical_spacing=0.03, subplot_titles=days_of_week_ordered, specs=[[{"secondary_y": True}]] * 7)
                        for i, day in enumerate(days_of_week_ordered):
                            if day in req_pivot.columns:
                                fig.add_trace(go.Bar(x=intervals_str_fmt, y=volume_pivot[day], name='Volume', marker_color='lightblue'), row=i+1, col=1, secondary_y=True)
                                fig.add_trace(go.Scatter(x=intervals_str_fmt, y=req_pivot[day], mode='lines', name='Required Staff', line=dict(color='crimson')), row=i+1, col=1, secondary_y=False)
                        fig.update_layout(height=1400, showlegend=False, title_text="Daily Volume and Required Staff")
                        fig.update_yaxes(title_text="<b>Required Staff</b>", secondary_y=False)
                        fig.update_yaxes(title_text="<b>Volume</b>", secondary_y=True, showgrid=False)
                        st.plotly_chart(fig, use_container_width=True)

                        st.markdown("##### Daily Staffing Requirements by Interval")
                        st.dataframe(req_pivot.style.background_gradient(cmap='viridis'), use_container_width=True)
                        download_dataframe_csv(req_pivot, f"staffing_requirements_{selected_scenario_for_detail}_{selected_week_for_detail}")

                        with st.expander("Peak Interval Analysis (Pareto)", expanded=False):
                            st.info("Identify the top intervals contributing to a certain percentage of the daily or weekly workload.")
                            pareto_threshold = st.slider("Pareto Threshold (%)", 1, 100, 80, key="pareto_slider")

                            pareto_col1, pareto_col2 = st.columns(2)
                            weekly_df_copy = weekly_df.copy()

                            with pareto_col1:
                                st.markdown("##### 📈 Top Volume Intervals")
                                weekly_df_copy['Week'] = weekly_df_copy['Week_Start_Day'].dt.strftime('%Y-%m-%d')
                                if weekly_df_copy['Volume'].sum() > 0:
                                    weekly_volume_pareto = pareto_analysis(weekly_df_copy, 'Volume', 'Week', pareto_threshold)
                                    st.write("Weekly")
                                    st.dataframe(weekly_volume_pareto.drop(columns=['Week', 'CumulativePercentage']).reset_index(drop=True).style.format({'Contribution (%)': '{:.2f}%'}))
                                    download_dataframe_csv_no_index(weekly_volume_pareto.drop(columns=['Week', 'CumulativePercentage']).reset_index(drop=True), "weekly_volume_pareto")
                                
                                    # Weekly Volume Pareto Chart
                                    st.markdown("##### Weekly Volume Distribution with Pareto Highlights")
                                    if not weekly_df_copy.empty and 'Volume' in weekly_df_copy.columns:
                                        # Group by Interval and sum Volume across all days of the week
                                        weekly_total_volume_by_interval = weekly_df_copy.groupby('Interval')['Volume'].sum().reset_index()
                                        weekly_total_volume_by_interval['IntervalStr'] = weekly_total_volume_by_interval['Interval'].apply(lambda t: t.strftime('%H:%M'))
                                        
                                        # Get intervals that are part of the Pareto for highlighting
                                        pareto_volume_intervals = weekly_volume_pareto['Interval'].tolist()
                                        weekly_total_volume_by_interval['color'] = np.where(weekly_total_volume_by_interval['IntervalStr'].isin(pareto_volume_intervals), 'orange', 'blue')
                                        
                                        fig_weekly_volume = go.Figure()
                                        fig_weekly_volume.add_trace(go.Bar(
                                            x=weekly_total_volume_by_interval['IntervalStr'],
                                            y=weekly_total_volume_by_interval['Volume'],
                                            marker_color=weekly_total_volume_by_interval['color'],
                                            name='Volume'
                                        ))
                                        fig_weekly_volume.update_layout(
                                            title="Weekly Volume Distribution with Pareto Highlights",
                                            xaxis_title="Time Interval",
                                            yaxis_title="Total Weekly Volume",
                                            showlegend=False
                                        )
                                        st.plotly_chart(fig_weekly_volume, use_container_width=True)
                                    else:
                                       st.info("No weekly volume data to display Pareto chart.")

                                pareto_day_vol = st.selectbox("Select Day for Volume Pareto", options=days_of_week_ordered, key="pareto_day_vol")
                                day_df_vol = weekly_df[weekly_df['Day'] == pareto_day_vol]
                                if not day_df_vol.empty and day_df_vol['Volume'].sum() > 0:
                                    daily_volume_pareto = pareto_analysis(day_df_vol, 'Volume', 'Day', pareto_threshold)
                                    st.dataframe(daily_volume_pareto.drop(columns=['Day', 'CumulativePercentage']).reset_index(drop=True).style.format({'Contribution (%)': '{:.2f}%'}))
                                    download_dataframe_csv_no_index(daily_volume_pareto.drop(columns=['Day', 'CumulativePercentage']).reset_index(drop=True), f"{pareto_day_vol}_volume_pareto")
                                    # Re-implemented Bar chart for Top Volume Intervals
                                    day_df_for_chart = day_df_vol.copy()
                                    pareto_intervals = daily_volume_pareto['Interval'].tolist()
                                    day_df_for_chart['IntervalStr'] = day_df_for_chart['Interval'].apply(lambda t: t.strftime('%H:%M'))
                                    day_df_for_chart['color'] = np.where(day_df_for_chart['IntervalStr'].isin(pareto_intervals), 'orange', 'blue')
                                    fig = go.Figure()
                                    fig.add_trace(go.Bar(
                                        x=day_df_for_chart['IntervalStr'],
                                        y=day_df_for_chart['Volume'],
                                        marker_color=day_df_for_chart['color'],
                                        name='Volume'
                                    ))
                                    fig.update_layout(
                                        title=f"Volume Analysis for {pareto_day_vol}",
                                        xaxis_title="Time Interval",
                                        yaxis_title="Volume",
                                        showlegend=False
                                    )
                                    st.plotly_chart(fig, use_container_width=True)
                                else:
                                    st.info("No volume data for this day to perform Pareto analysis.")
                            with pareto_col2:
                                st.markdown("##### 📈 Top Required Staff Intervals")
                                if weekly_df_copy['final_positions'].sum() > 0:
                                    weekly_req_pareto = pareto_analysis(weekly_df_copy, 'final_positions', 'Week', pareto_threshold)
                                    st.write("Weekly")
                                    st.dataframe(weekly_req_pareto.drop(columns=['Week', 'CumulativePercentage']).reset_index(drop=True).style.format({'Contribution (%)': '{:.2f}%'}))
                                    download_dataframe_csv_no_index(weekly_req_pareto.drop(columns=['Week', 'CumulativePercentage']).reset_index(drop=True), "weekly_req_pareto")
                                
                                    # Weekly Required Staff Pareto Chart
                                    st.markdown("##### Weekly Required Staff Distribution with Pareto Highlights")
                                    if not weekly_df_copy.empty and 'final_positions' in weekly_df_copy.columns:
                                        # Group by Interval and sum final_positions across all days of the week
                                        weekly_total_req_by_interval = weekly_df_copy.groupby('Interval')['final_positions'].sum().reset_index()
                                        weekly_total_req_by_interval['IntervalStr'] = weekly_total_req_by_interval['Interval'].apply(lambda t: t.strftime('%H:%M'))
                                        
                                        # Get intervals that are part of the Pareto for highlighting
                                        pareto_req_intervals = weekly_req_pareto['Interval'].tolist()
                                        weekly_total_req_by_interval['color'] = np.where(weekly_total_req_by_interval['IntervalStr'].isin(pareto_req_intervals), 'orange', 'blue')
                                        
                                        fig_weekly_req = go.Figure()
                                        fig_weekly_req.add_trace(go.Bar(
                                            x=weekly_total_req_by_interval['IntervalStr'],
                                            y=weekly_total_req_by_interval['final_positions'],
                                            marker_color=weekly_total_req_by_interval['color'],
                                            name='Required Staff'
                                        ))
                                        fig_weekly_req.update_layout(
                                            title="Weekly Required Staff Distribution with Pareto Highlights",
                                            xaxis_title="Time Interval",
                                            yaxis_title="Total Weekly Required Staff",
                                            showlegend=False
                                        )
                                        st.plotly_chart(fig_weekly_req, use_container_width=True)
                                    else:
                                       st.info("No weekly required staff data to display Pareto chart.")

                                else:
                                    st.info("No required staff data for this week to perform Pareto analysis.")
                                pareto_day_req = st.selectbox("Select Day for Requirement Pareto", options=days_of_week_ordered, key="pareto_day_req")
                                day_df_req = weekly_df[weekly_df['Day'] == pareto_day_req]
                                if not day_df_req.empty and day_df_req['final_positions'].sum() > 0:
                                    daily_req_pareto = pareto_analysis(day_df_req, 'final_positions', 'Day', pareto_threshold)
                                    st.dataframe(daily_req_pareto.drop(columns=['Day', 'CumulativePercentage']).reset_index(drop=True).style.format({'Contribution (%)': '{:.2f}%'}))
                                    download_dataframe_csv_no_index(daily_req_pareto.drop(columns=['Day', 'CumulativePercentage']).reset_index(drop=True), f"{pareto_day_req}_req_pareto")
                                    
                                    # Re-implemented Bar chart for Top Required Staff Intervals
                                    day_df_for_chart = day_df_req.copy()
                                    pareto_intervals = daily_req_pareto['Interval'].tolist()
                                    day_df_for_chart['IntervalStr'] = day_df_for_chart['Interval'].apply(lambda t: t.strftime('%H:%M'))
                                    day_df_for_chart['color'] = np.where(day_df_for_chart['IntervalStr'].isin(pareto_intervals), 'orange', 'blue')
                                    fig = go.Figure()
                                    fig.add_trace(go.Bar(
                                        x=day_df_for_chart['IntervalStr'],
                                        y=day_df_for_chart['final_positions'],
                                        marker_color=day_df_for_chart['color'],
                                        name='Required Staff'
                                    ))
                                    fig.update_layout(
                                        title=f"Required Staff Analysis for {pareto_day_req}",
                                        xaxis_title="Time Interval",
                                        yaxis_title="Required Staff",
                                        showlegend=False
                                    )
                                    st.plotly_chart(fig, use_container_width=True)
                                else:
                                    st.info("No required staff data for this day to perform Pareto analysis.")

                    with tab_data_details:
                        st.subheader(f"Detailed KPI Summary for Week of {selected_week_for_detail}")

                        is_erlang_scenario = 'service_level' in weekly_df.columns

                        if is_erlang_scenario:
                            st.markdown("##### Weekly Summary")
                            weekly_kpis = calculate_aggregated_kpis(weekly_df)
                            kpi_cols = st.columns(7)
                            kpi_cols[0].metric("Total Calls", f"{weekly_kpis['Total Calls']:,}")
                            kpi_cols[1].metric("Total Raw Positions", f"{weekly_kpis['Total Raw Positions']:,}")
                            kpi_cols[2].metric("Total Final Positions", f"{weekly_kpis['Total Final Positions']:,}")
                            kpi_cols[3].metric("Weighted Overall ASA (s)", f"{weekly_kpis['Overall ASA (s)']:.2f}s")
                            kpi_cols[4].metric("Weighted Service Level", f"{weekly_kpis['Service Level (%)']:.2f}%")
                            kpi_cols[5].metric("Weighted Occupancy", f"{weekly_kpis['Occupancy (%)']:.2f}%")
                            kpi_cols[6].metric("Weighted Wait Prob.", f"{weekly_kpis['Wait Probability (%)']:.2f}%")

                            # Fetch and display the weekly HC metrics first
                            summary_row = st.session_state.scenario_summary[
                                (st.session_state.scenario_summary['Scenario'] == selected_scenario_for_detail) &
                                (st.session_state.scenario_summary['Week_Start_Day'] == selected_week_for_detail)
                            ]
                            if not summary_row.empty:
                                hc_avg = summary_row.iloc[0]['Required HC (Avg)']
                                hc_peak = summary_row.iloc[0]['Required HC for Peak Day']
                                hc_cols = st.columns(2)
                                hc_cols[0].metric("Required HC (Avg)", f"{hc_avg:.2f}")
                                hc_cols[1].metric("Required HC for Peak Day", f"{hc_peak:.2f}")


                            st.markdown("##### Daily Volume-Weighted KPIs")
                            daily_kpi_rows = []
                            for day in days_of_week_ordered:
                                daily_df = weekly_df[weekly_df['Day'] == day]
                                if not daily_df.empty:
                                    daily_kpis = calculate_aggregated_kpis(daily_df)
                                    daily_kpis['Day'] = day
                                    daily_kpi_rows.append(daily_kpis)

                            if daily_kpi_rows:
                                daily_kpi_df = pd.DataFrame(daily_kpi_rows).set_index("Day")
                                st.dataframe(daily_kpi_df[[
                                    "Total Calls", "Total Raw Positions", "Total Final Positions",
                                    "Required HC per day", "Required Raw per day", # New Columns
                                    "Service Level (%)", "Occupancy (%)", "Wait Probability (%)", "Overall ASA (s)"
                                ]].style.format({
                                    'Total Calls': '{:,.0f}', 'Total Raw Positions': '{:,.0f}', 'Total Final Positions': '{:,.0f}',
                                    'Required HC per day': '{:.2f}', 'Required Raw per day': '{:.2f}', # Formatting
                                    'Service Level (%)': '{:.2f}%', 'Occupancy (%)': '{:.2f}%',
                                    'Wait Probability (%)': '{:.2f}%', 'Overall ASA (s)': '{:.2f}s'
                                }), use_container_width=True)
                                download_dataframe_csv(daily_kpi_df, f"daily_kpis_{selected_scenario_for_detail}_{selected_week_for_detail}")

                        else:
                            st.info("KPIs like Service Level and ASA are only applicable for Erlang-based scenarios (Voice and Chat).")

                        st.markdown("---")
                        st.markdown("##### Interval Level Data")

                        cols_to_show = ['Date', 'Day', 'Interval', 'Volume', 'raw_positions', 'final_positions']
                        # Retain other columns for Erlang scenarios, but exclude problematic ASA/AWT
                        if is_erlang_scenario:
                             cols_to_show.extend(['service_level', 'occupancy', 'waiting_probability'])

                        interval_detail_df = weekly_df.reindex(columns=cols_to_show).copy()
                        interval_detail_df['Interval'] = interval_detail_df['Interval'].apply(lambda t: t.strftime('%H:%M'))

                        format_dict = {
                            'raw_positions': '{:.0f}', 'final_positions': '{:.0f}', 'Volume': '{:.0f}'
                        }
                        if is_erlang_scenario:
                            format_dict.update({
                                'service_level': '{:.2%}', 'occupancy': '{:.2%}',
                                'waiting_probability': '{:.2%}'
                            })

                        st.dataframe(interval_detail_df.style.format(format_dict, na_rep="-"), height=500, use_container_width=True)
                        download_dataframe_csv_no_index(interval_detail_df, f"interval_detail_{selected_scenario_for_detail}_{selected_week_for_detail}")

with tab2:
    st.header("Step 2: Generate & Cost Schedule Shells")
    st.info("Define shifts and scheduling rules, then generate the most efficient weekly roster to meet demand. **Use the new Pre-flight Analyzer to check for issues before solving!**")

    st.sidebar.header("📜 Roster & Costing Rules")
    with st.sidebar.expander("💰 Pay & Overtime Rules", expanded=True):
        st.number_input("Base Hourly Rate ($)", min_value=10.0, value=st.session_state.get('base_hourly_rate', 20.0), step=0.5, key="base_hourly_rate")
        st.number_input("Weekly Overtime Threshold (hours)", min_value=0, value=st.session_state.get('ot_hours_threshold', 40), key="ot_hours_threshold")
        st.number_input("Overtime Rate Multiplier", min_value=1.0, value=st.session_state.get('ot_rate_multiplier', 1.5), step=0.1, key="ot_rate_multiplier")
    
    sched_mode_options = ("Use Pre-defined Shifts", "Optimize Shifts Automatically")
    saved_sched_mode = st.session_state.get('sched_mode', sched_mode_options[0])
    try:
        sched_mode_index = sched_mode_options.index(saved_sched_mode)
    except ValueError:
        sched_mode_index = 0
    schedule_generation_mode = st.sidebar.radio(
        "Select Scheduling Mode",
        sched_mode_options,
        index=sched_mode_index,
        key="sched_mode",
        help="Choose 'Pre-defined' to test your own shifts. Choose 'Optimize' to have the solver design the best possible shifts for you."
    )

    if schedule_generation_mode == "Use Pre-defined Shifts":
        with st.sidebar.expander("📅 Shift Pattern Definitions", expanded=True):
            st.markdown("**Define Shift Timings**")
            st.info("Define shifts by start time and total length. All shift hours are considered paid.")
            edited_shifts_df = st.data_editor(
                st.session_state.shifts_df,
                num_rows="dynamic",
                key="shifts_editor",
                column_config={
                    "Shift Name": st.column_config.TextColumn(required=True),
                    "Start Time": st.column_config.TimeColumn(format="HH:mm", required=True),
                    "Shift Length (hours)": st.column_config.NumberColumn(min_value=1.0, max_value=16.0, step=0.5, required=True, help="Total duration of the shift from start to end."),
                },
                use_container_width=True,
            )
            st.session_state.shifts_df = edited_shifts_df.copy()
            download_dataframe_csv_no_index(st.session_state.shifts_df, "defined_shifts")
    else: # Advanced Shift Optimization Mode
        with st.sidebar.expander("🤖 Advanced Shift Optimization Rules", expanded=True):
            st.markdown("##### 1. Select Optimization Model")

            opt_model_options = ["Best Fit (Balanced)", "Line Adherence (Coverage Target)"]
            default_opt_model = st.session_state.get('optimization_model_choice', "Best Fit (Balanced)")
            try:
                default_opt_idx = opt_model_options.index(default_opt_model)
            except ValueError:
                default_opt_idx = 0

            optimization_model_choice = st.radio(
                "Select Optimization Model",
                options=opt_model_options,
                index=default_opt_idx,
                key="optimization_model_choice",
                horizontal=True,
                label_visibility="collapsed"
            )

            if optimization_model_choice == "Best Fit (Balanced)":
                st.markdown("##### 2. Optimization Objectives")
                st.info("The primary goal is to maximize coverage by minimizing understaffing, weighted by how high the demand is. The penalties below fine-tune this behavior.")
                st.slider("Understaffing Penalty Weight", 1, 100, value=st.session_state.get('understaff_penalty_opt', 15), key="understaff_penalty_opt", help="Higher values make the solver prioritize covering every required slot, even if it causes overstaffing.")
                st.slider("Overstaffing Penalty Weight", 1, 100, value=st.session_state.get('overstaff_penalty_opt', 1), key="overstaff_penalty_opt", help="How much to penalize surplus staff.")
            else: # Line Adherence Model
                st.markdown("##### 2. Adherence Model Settings")
                st.info("The goal is to meet the adherence target with the minimum possible staff. The solver will find the cheapest roster that satisfies the target.")
                
                opt_adherence_level_options = ["Day", "Week"]
                saved_opt_adherence_level = st.session_state.get('opt_adherence_target_level', "Day")
                try:
                    opt_adherence_level_index = opt_adherence_level_options.index(saved_opt_adherence_level)
                except ValueError:
                    opt_adherence_level_index = 0
                st.radio("Target Period", opt_adherence_level_options, index=opt_adherence_level_index, horizontal=True, key="opt_adherence_target_level")
                
                st.slider("Line Adherence Target (%)", 80, 120, value=st.session_state.get('opt_adherence_target_percent', 95), key="opt_adherence_target_percent")
                st.slider("Interval Overstaffing Cap (%)", 100, 200, value=st.session_state.get('opt_adherence_cap_percent', 105), key="opt_adherence_cap_percent")

            st.markdown("##### 3. Daily Operational Hours")
            st.info(
                "Set the daily start and end times within which all optimized shifts must fall. "
                "**For 24/7 operations on a specific day, set both Start and End Time to 00:00.**"
            )
            if st.button("Set All Days to 24/7 Operation", key="set_24_7_op_hours"):
                for day in DAYS_OF_WEEK_OPTIONS:
                    st.session_state.daily_op_hours[day] = {"Start Time": datetime.time(0, 0), "End Time": datetime.time(0, 0)}
                st.rerun()

            op_hours_df = pd.DataFrame.from_dict(st.session_state.daily_op_hours, orient='index')
            edited_op_hours_df = st.data_editor(
                op_hours_df,
                key="daily_op_hours_editor",
                column_config={
                    "Start Time": st.column_config.TimeColumn(format="HH:mm", required=True),
                    "End Time": st.column_config.TimeColumn(format="HH:mm", required=True),
                },
                use_container_width=True
            )
            st.session_state.daily_op_hours = edited_op_hours_df.to_dict(orient='index')


            st.markdown("##### 4. Shift Creation Parameters")
            st.checkbox("Keep employees on same shift for entire week",
                        value=st.session_state.get('shift_consistency_opt', True),
                        key='shift_consistency_opt',
                        help="When enabled, each employee maintains a consistent shift timing (e.g., 8am-5pm) throughout the week.")
            st.number_input("Maximum Unique Shifts to Create", min_value=1, max_value=20, value=st.session_state.get('max_unique_shifts', 5), key="max_unique_shifts", help="Limits how many different shift patterns the system can generate.")

            st.markdown("##### 5. Work Rules per Shift Length")
            duration_options = [i / 2.0 for i in range(1, 25)] # 0.5 to 12.0
            st.multiselect("Allowed Shift Durations (hours)", options=duration_options, key='allowed_durations', default=st.session_state.get('allowed_durations', [8.0, 10.0]))

            current_rules = st.session_state.get('duration_rules', {})
            st.session_state.duration_rules = {dur: rule for dur, rule in current_rules.items() if dur in st.session_state.allowed_durations}
            current_caps = st.session_state.get('distribution_caps', {})
            st.session_state.distribution_caps = {dur: cap for dur, cap in current_caps.items() if dur in st.session_state.allowed_durations}

            if not st.session_state.allowed_durations:
                st.warning("Please select at least one shift duration.")
            else:
                for dur in sorted(st.session_state.allowed_durations):
                    if dur not in st.session_state.duration_rules:
                        default_days = 5 if dur <= 8.0 else 4
                        st.session_state.duration_rules[dur] = {'min_days': default_days, 'max_days': default_days, 'min_off': 2}

                    cols = st.columns(2)
                    work_day_range = cols[0].slider(
                        f"Work Days/Wk for {dur}hr shifts",
                        min_value=1, max_value=7,
                        value=(
                            st.session_state.duration_rules[dur].get('min_days', 4),
                            st.session_state.duration_rules[dur].get('max_days', 5)
                        ),
                        key=f'days_range_for_{dur}hr'
                    )
                    st.session_state.duration_rules[dur]['min_days'] = work_day_range[0]
                    st.session_state.duration_rules[dur]['max_days'] = work_day_range[1]

                    st.session_state.duration_rules[dur]['min_off'] = cols[1].number_input(f"Min Consecutive Off for {dur}hr shifts", 1, 4, value=st.session_state.duration_rules[dur].get('min_off', 2), key=f'off_for_{dur}hr')

            st.markdown("##### 6. Staffing & Distribution Controls (Optional)")
            sc1, sc2 = st.columns(2)
            st.number_input("Min Agents per Shift Type", min_value=0, value=st.session_state.get('min_agents_per_shift', 0), key='min_agents_per_shift', help="If a shift type is used, it must have at least this many total weekly assignments. Set to 0 to disable.")
            st.number_input("Max Agents per Shift Type", min_value=0, value=st.session_state.get('max_agents_per_shift', 100), key='max_agents_per_shift', help="If a shift type is used, it can have at most this many total weekly assignments. Set to a high number to disable.")

            st.markdown("###### Shift Distribution Caps")
            for dur in sorted(st.session_state.allowed_durations):
                if dur not in st.session_state.distribution_caps:
                     st.session_state.distribution_caps[dur] = {'enabled': False, 'percent': 30}

                with st.container(border=True):
                    st.session_state.distribution_caps[dur]['enabled'] = st.checkbox(f"Cap {dur}hr Shifts", key=f'cap_enabled_{dur}', value=st.session_state.distribution_caps[dur].get('enabled', False))
                    if st.session_state.distribution_caps[dur]['enabled']:
                        st.session_state.distribution_caps[dur]['percent'] = st.slider(f"Max % for {dur}hr Shifts", 0, 100, value=st.session_state.distribution_caps[dur].get('percent', 30), key=f'cap_percent_{dur}')

            st.markdown("##### 7. Headcount for Solver")
            st.number_input("Total Headcount to Schedule", min_value=0, value=st.session_state.get('total_hc_optimization', 20), key='total_hc_optimization', help="The total number of employees the optimizer can use.")

    with st.sidebar.expander("💸 Shift & Holiday Differentials", expanded=False):
        st.markdown("**Shift Differentials (by time of day)**")
        if 'shift_differentials_df' not in st.session_state:
            st.session_state.shift_differentials_df = pd.DataFrame([
                {"Name": "Evening Premium", "Start Time": datetime.time(18, 0), "End Time": datetime.time(23, 0), "Premium Type": "Additive ($)", "Premium": 2.50},
                {"Name": "Night Owl", "Start Time": datetime.time(23, 0), "End Time": datetime.time(6, 0), "Premium Type": "Percentage", "Premium": 15.0}])

        st.session_state.shift_differentials_df = st.data_editor(st.session_state.shift_differentials_df, num_rows="dynamic", key='shift_diff_editor',
                                                column_config={"Start Time": st.column_config.TimeColumn(format="HH:mm"), "End Time": st.column_config.TimeColumn(format="HH:mm"),
                                                               "Premium Type": st.column_config.SelectboxColumn(options=["Additive ($)", "Percentage"])})
        download_dataframe_csv_no_index(st.session_state.shift_differentials_df, "shift_differentials")

        st.markdown("**Weekend & Holiday Differentials**")
        all_dates_in_view = pd.date_range(start=start_date, end=end_date)
        st.multiselect("Select Public Holidays", options=all_dates_in_view.date, default=st.session_state.get('holiday_dates', []), format_func=lambda d: d.strftime('%Y-%m-%d (%A)'), key="holiday_dates")
        st.text_input("Holiday Premium Name", value=st.session_state.get('holiday_prem_name', "Public Holiday Pay"), key="holiday_prem_name")
        st.number_input("Holiday Premium Multiplier", min_value=1.0, value=st.session_state.get('holiday_prem_mult', 2.0), key="holiday_prem_mult")
        sunday_pay = st.checkbox("Apply Sunday Premium", value=st.session_state.get('sunday_pay_check', True), key="sunday_pay_check")
        st.number_input("Sunday Premium Multiplier", min_value=1.0, value=st.session_state.get('sunday_prem_mult', 1.5), disabled=not sunday_pay, key="sunday_prem_mult")

    with st.sidebar.expander("⚖️ General Scheduling Constraints", expanded=False):
        if schedule_generation_mode == "Use Pre-defined Shifts":
            st.markdown("**Work Days per Week (by Shift)**")
            st.info("Set the number of working days per week for employees assigned to each shift type.")
            work_days_by_shift = {}
            defined_shifts = st.session_state.shifts_df['Shift Name'].dropna().unique()
            if len(defined_shifts) > 0:
                for shift_name in defined_shifts:
                    work_days_by_shift[shift_name] = st.slider(
                        f"Days/wk for '{shift_name}' shift", 1, 7, st.session_state.get(f"work_days_{sanitize_name(shift_name)}", 5),
                        key=f"work_days_{sanitize_name(shift_name)}"
                    )
            else:
                st.warning("Define at least one shift to set workday rules.")
            st.checkbox("Enforce Same Shift for Entire Week", value=True, disabled=True, help="This is required to apply different workday rules per shift type and is always active.")
            st.markdown("---")
        else:
            work_days_by_shift = {}

        st.slider("Max consecutive work days", 4, 7, value=st.session_state.get('max_consecutive_slider', 6), key="max_consecutive_slider")
        st.slider("Min consecutive days off (for Pre-defined shifts)", 1, 3, value=st.session_state.get('min_off_days', 2), help="For 'Pre-defined shifts' mode only. For Optimization mode, this is set per shift duration.", key="min_off_days", disabled=(schedule_generation_mode != "Use Pre-defined Shifts"))

        st.markdown("---")
        st.checkbox(
            "Attempt to schedule with insufficient headcount",
            value=st.session_state.get('force_schedule_insufficient_hc', False),
            key="force_schedule_insufficient_hc",
            help="If checked, the scheduler will run only the relaxed 'Best Fit' model with relaxed work-life-balance rules to find the best possible coverage, even if it results in understaffing."
        )


    with st.sidebar.expander("🎯 'Line Adherence' Model Settings (Pre-defined shifts only)", expanded=False):
        st.info("Only applicable when using 'Use Pre-defined Shifts' mode.")
        enable_adherence_model = st.checkbox("Enable Line Adherence Scheduling Model", value=st.session_state.get('enable_adherence', False), key="enable_adherence", disabled=(schedule_generation_mode != "Use Pre-defined Shifts"))

        adherence_level_options = ["Day", "Week"]
        default_adherence_level = st.session_state.get('adherence_target_level', 'Day')
        try:
            default_adherence_idx = adherence_level_options.index(default_adherence_level)
        except ValueError:
            default_adherence_idx = 0

        if enable_adherence_model:
            st.radio("Target Level", options=adherence_level_options, index=default_adherence_idx, horizontal=True, help="Choose 'Day' for consistent daily adherence. Choose 'Week' for flexibility.", key="adherence_target_level")
            st.slider("Line Adherence Target (%)", 80, 120, value=st.session_state.get('adherence_target_percent', 95), key="adherence_target_percent")
            st.slider("Interval Overstaffing Cap (%)", 100, 200, value=st.session_state.get('adherence_cap_percent', 105), key="adherence_cap_percent")

    st.markdown("#### 1. Select Requirement Input Source")
    input_source_options = ("Use Staffing Forecast from Tab 1", "Manually Enter Requirements")
    saved_input_source = st.session_state.get('t2_input_source', input_source_options[0])
    try:
        input_source_index = input_source_options.index(saved_input_source)
    except ValueError:
        input_source_index = 0
    input_source = st.radio(
        "Where should the staffing requirements come from?",
        input_source_options,
        index=input_source_index,
        key="t2_input_source", label_visibility="collapsed"
    )

    all_available_jobs = []
    if input_source == "Use Staffing Forecast from Tab 1":
        if "scenario_summary" not in st.session_state or st.session_state.scenario_summary.empty:
            st.warning("Please run a Staffing Calculation on Tab 1 first to generate a forecast.", icon="⚠️")
        else:
            summary_df = st.session_state.scenario_summary
            available_scenarios = summary_df['Scenario'].unique()
            selected_scenarios = st.multiselect("Select Forecast Scenario(s) to Schedule", available_scenarios, default=st.session_state.get('t2_selected_scenarios', available_scenarios), key='t2_selected_scenarios')

            if selected_scenarios:
                weeks_df = summary_df[summary_df['Scenario'].isin(selected_scenarios)][['Scenario', 'Week_Start_Day', 'Required HC (Avg)']].drop_duplicates()
                weeks_df['display'] = weeks_df.apply(lambda row: f"{row['Scenario']} | Week of {row['Week_Start_Day']} | HC: {row['Required HC (Avg)']:.1f}", axis=1)
                all_available_jobs.extend(weeks_df.to_dict('records'))

    else: # Manual Input
        st.markdown("#### Define Manual Requirements")
        manual_cols = st.columns(2)
        with manual_cols[0]:
            manual_start_date = st.date_input("Start Date", value=st.session_state.get('manual_start_date', datetime.date.today()), key="manual_start_date")
        with manual_cols[1]:
            manual_end_date = st.date_input("End Date", value=st.session_state.get('manual_end_date', datetime.date.today() + datetime.timedelta(days=6)), key="manual_end_date")

        if manual_start_date > manual_end_date:
            st.error("Error: End date must be after start date.")
        else:
            manual_date_range = pd.date_range(manual_start_date, manual_end_date)
            manual_dates_str = manual_date_range.strftime('%Y-%m-%d').tolist()
            interval_index = [t.strftime('%H:%M:%S') for t in st.session_state.intervals]

            # Apply safe update to manual requirements dataframe
            if "manual_req_df" not in st.session_state or list(st.session_state["manual_req_df"].columns) != manual_dates_str:
                safe_update_dataframe("manual_req_df", manual_dates_str, interval_index)

            st.info("Enter the number of required staff for each interval. Columns are dates.")
            st.session_state["manual_req_df"] = st.data_editor(st.session_state["manual_req_df"], key="manual_req_editor", height=350, use_container_width=True)
            download_dataframe_csv(st.session_state["manual_req_df"], "manual_requirements")

            processed_manual_weeks = process_manual_requirements(
                st.session_state["manual_req_df"],
                st.session_state.week_start_day,
                st.session_state.working_days,
                st.session_state.working_hours,
                days_of_week_ordered
            )
            if processed_manual_weeks:
                all_available_jobs.extend(processed_manual_weeks)

    # --- UNIFIED SELECTION LOGIC ---
    jobs_for_preview = []
    if all_available_jobs:
        st.markdown("##### Select which jobs to generate schedules for:")
        run_mode_options = ("Run for ALL available jobs listed below", "Select SPECIFIC jobs to run")
        saved_run_mode = st.session_state.get('scheduling_run_mode', run_mode_options[0])
        try:
            run_mode_index = run_mode_options.index(saved_run_mode)
        except ValueError:
            run_mode_index = 0
        run_mode = st.radio(
            "Scheduling Run Mode",
            run_mode_options,
            index=run_mode_index,
            key="scheduling_run_mode",
            label_visibility="collapsed"
        )

        if run_mode == "Run for ALL available jobs listed below":
            jobs_for_preview = all_available_jobs
            with st.expander("Jobs to be processed", expanded=False):
                for job in jobs_for_preview:
                    st.write(f"- {job['display']}")
        else: # "Select SPECIFIC jobs to run"
            job_options = [job['display'] for job in all_available_jobs]
            selected_jobs_display = st.multiselect(
                "Select jobs to run:",
                options=job_options,
                default=st.session_state.get('t2_selected_jobs', []) if 't2_selected_jobs' in st.session_state else job_options,
                key='t2_selected_jobs'
            )
            jobs_for_preview = [job for job in all_available_jobs if job['display'] in selected_jobs_display]


    st.markdown("---")
    st.markdown("#### 2. In-Office Shrinkage (Optional)")
    st.info("Set a non-productive time percentage (e.g., for coaching, meetings) for each interval and day. This will increase the staff requirement for the scheduler. This is separate from the shrinkage in Tab 1. Enter values like '15' for 15%.")

    st.markdown("##### Apply Shrinkage Pattern")
    st.caption("Optional: Set shrinkage for one day, then apply that pattern to other days to save time.")
    pattern_cols = st.columns([1, 2, 1])
    with pattern_cols[0]:
        source_day = st.selectbox("Source Day", options=days_of_week_ordered, key="shrinkage_source_day")
    with pattern_cols[1]:
        target_days = st.multiselect("Target Day(s)", options=[d for d in days_of_week_ordered if d != source_day], key="shrinkage_target_days")
    with pattern_cols[2]:
        st.write("")
        st.write("")
        if st.button("Apply Pattern", key="apply_shrinkage_pattern"):
            if source_day and target_days:
                source_pattern = st.session_state.daily_shrinkage_df[source_day].copy()
                for day in target_days:
                    st.session_state.daily_shrinkage_df[day] = source_pattern
                st.success(f"Applied {source_day}'s pattern to {', '.join(target_days)}.")
                st.rerun()
            else:
                st.warning("Please select a source day and at least one target day.")

    st.session_state.daily_shrinkage_df = st.data_editor(
        st.session_state.daily_shrinkage_df,
        column_config={col: st.column_config.NumberColumn(label=f"{col} Shrinkage %", min_value=0.0, max_value=99.9, step=0.5, format="%.1f%%") for col in days_of_week_ordered},
        key="daily_shrinkage_editor",
        use_container_width=True,
        height=350
    )
    download_dataframe_csv(st.session_state.daily_shrinkage_df, "daily_shrinkage_matrix")

    st.markdown("##### Daily Shrinkage Pattern Chart")
    st.info("Visualizes the shrinkage percentages applied across different days and intervals.")
    shrinkage_df_for_plot = st.session_state.daily_shrinkage_df.copy()

    fig_shrinkage = go.Figure(data=go.Heatmap(
        z=shrinkage_df_for_plot.values.T,
        x=shrinkage_df_for_plot.index,
        y=shrinkage_df_for_plot.columns,
        colorscale='Viridis',
        hovertemplate='Day: %{y}<br>Interval: %{x}<br>Shrinkage: %{z:.1f}%<extra></extra>'
    ))
    fig_shrinkage.update_layout(
        title='Daily Shrinkage Pattern',
        xaxis_title='Time Interval',
        yaxis_title='Day of Week',
        yaxis=dict(autorange='reversed')
    )
    st.plotly_chart(fig_shrinkage, use_container_width=True)


    if jobs_for_preview:
        st.markdown("---")
        if schedule_generation_mode == "Use Pre-defined Shifts":
            st.markdown("#### 3. Headcount for Solver (Final Adjustment)")
            st.info(
                "Below, you can override the calculated **Base HC** for each selected job. "
                "This final number will be used for all models (Best Fit, Meet or Exceed, etc.) "
                "to ensure a fair comparison. You can increase the number to add a buffer or "
                "decrease it to simulate staff shortages."
            )
            for job in jobs_for_preview:
                display_name = job.get('display')
                if not display_name:
                    week_start_dt = job.get('week_start_dt', 'N/A')
                    display_name = f"Manual Input | Week of {week_start_dt.strftime('%Y-%m-%d') if hasattr(week_start_dt, 'strftime') else week_start_dt}"

                base_hc = math.ceil(job.get('Required HC (Avg)', job.get('avg_fte', 0)))

                st.session_state.adjusted_headcounts[display_name] = st.number_input(
                    f"**Final number of employees for: {display_name.split(' | HC:')[0]}**",
                    min_value=1,
                    value=st.session_state.adjusted_headcounts.get(display_name, base_hc),
                    key=f"headcount_adjust_{sanitize_name(display_name)}",
                    help=f"The calculated Base HC for this week is {base_hc}. Adjust this number to simulate having more or fewer staff available for the schedule."
                )
        else:
            pass

    st.markdown("---")
    st.markdown("#### 4. Pre-flight Schedule Analysis")
    st.info("Before running the time-consuming solver, analyze your constraints against your requirements to catch impossible or difficult-to-solve scenarios.")

    if st.button("Analyze Feasibility", key="analyze_feasibility_button"):
        if not jobs_for_preview:
            st.warning("Please select at least one Forecast or Manual week to analyze.")
        else:
            with st.spinner("Analyzing constraints for selected jobs..."):
                for job in jobs_for_preview:
                    display_name = job.get('display')
                    if 'Scenario' in job:
                        scenario, week_start_str = job['Scenario'], job['Week_Start_Day']
                        week_start_dt = datetime.datetime.strptime(week_start_str, '%Y-%m-%d').date()
                        scenario_data = st.session_state.all_scenarios[scenario]
                        full_staffing_df = scenario_data[0]
                        week_dates = pd.to_datetime([(week_start_dt + datetime.timedelta(days=d)) for d in range(7)])
                        req_pivot = full_staffing_df.pivot_table(index='Date', columns='Interval', values='final_positions', fill_value=0).reindex(index=week_dates, fill_value=0)
                        base_req_matrix = np.ceil(req_pivot.values).astype(int).tolist()
                    else:
                        week_start_dt = job['week_start_dt']
                        display_name = f"Manual Input | Week of {week_start_dt.strftime('%Y-%m-%d')}"
                        base_req_matrix = job['matrix']

                    num_days, num_intervals = len(base_req_matrix), len(base_req_matrix[0]) if base_req_matrix else 0
                    inflated_req_matrix = [[0] * num_intervals for _ in range(num_days)]
                    for d in range(num_days):
                        day_name = days_of_week_ordered[d]
                        day_shrinkage = st.session_state.daily_shrinkage_df[day_name].tolist()
                        for p in range(num_intervals):
                            original_req = base_req_matrix[d][p]
                            shrinkage_percent = day_shrinkage[p]
                            denominator = 1 - (shrinkage_percent / 100.0)
                            inflated_req = math.ceil(original_req / denominator) if (original_req > 0 and denominator > 0) else original_req
                            inflated_req_matrix[d][p] = int(inflated_req)

                    analyzer_kwargs = {
                        'days_of_week_ordered': days_of_week_ordered,
                        'intervals': st.session_state.intervals,
                        'force_schedule_insufficient_hc': st.session_state.get('force_schedule_insufficient_hc', False)
                    }
                    if schedule_generation_mode == "Use Pre-defined Shifts":
                        headcount = st.session_state.adjusted_headcounts.get(display_name, math.ceil(job.get('Required HC (Avg)', job.get('avg_fte', 0))))
                        analyzer_kwargs['shifts_df'] = st.session_state.shifts_df
                        analyzer_kwargs['work_days_by_shift'] = work_days_by_shift
                    else:
                        headcount = st.session_state.total_hc_optimization
                        analyzer_kwargs['allowed_durations'] = st.session_state.allowed_durations
                        analyzer_kwargs['duration_rules'] = st.session_state.duration_rules
                        analyzer_kwargs['daily_op_hours'] = st.session_state.daily_op_hours

                    findings = analyze_schedule_feasibility(inflated_req_matrix, headcount, schedule_generation_mode, **analyzer_kwargs)

                    with st.expander(f"**Analysis for: {display_name.split(' | HC:')[0]}**", expanded=True):
                        if not findings:
                            st.success("✅ **Looks Good!** No obvious constraint conflicts found. The solver should be able to start.", icon="✅")
                        else:
                            severities = {f['severity'] for f in findings}
                            if 'CRITICAL' in severities:
                                st.error("🛑 **Infeasible!** Critical issues found that will likely cause the solver to fail. Please address them.", icon="🛑")
                            elif 'HIGH' in severities:
                                st.error("🔥 **Highly Challenging!** The current setup is very constrained and may be impossible to solve. Review the issues below.", icon="🔥")
                            else:
                                st.warning("🤔 **Potential Challenges Found.** The solver might struggle or produce a sub-optimal schedule. Review the suggestions below.", icon="🤔")

                            for finding in findings:
                                if finding['severity'] in ['CRITICAL', 'HIGH']:
                                    st.error(f"**{finding['type'].replace('_', ' ')}:** {finding['message']}")
                                else:
                                    st.warning(f"**{finding['type'].replace('_', ' ')}:** {finding['message']}")
                                st.info(f"💡 **Suggestion:** {finding['suggestion']}")

    is_ready_to_run = jobs_for_preview and (
        schedule_generation_mode == "Use Pre-defined Shifts" or
        (schedule_generation_mode == "Optimize Shifts Automatically" and st.session_state.allowed_durations)
    )

    st.markdown("---")
    st.markdown("#### 5. Generate Schedules")
    if st.button("Generate & Cost Schedules", disabled=not is_ready_to_run, type="primary"):
        start_time = time.time()
        constraints_config = {'max_consecutive_work': st.session_state.max_consecutive_slider, 'min_consecutive_off': st.session_state.min_off_days}
        day_differentials = {'Sunday': {'name': 'Sunday Pay', 'type': 'Multiplier', 'value': st.session_state.sunday_prem_mult}} if st.session_state.sunday_pay_check else {}
        holiday_differentials = {d.strftime('%Y-%m-%d'): {'name': st.session_state.holiday_prem_name, 'type': 'Multiplier', 'value': st.session_state.holiday_prem_mult} for d in st.session_state.get('holiday_dates', [])}
        cost_config = {'base_rate': st.session_state.base_hourly_rate, 'ot_threshold': st.session_state.ot_hours_threshold, 'ot_multiplier': st.session_state.ot_rate_multiplier, 'shift_differentials': st.session_state.shift_differentials_df, 'day_differentials': day_differentials, 'holidays': holiday_differentials}
        force_schedule_flag = st.session_state.get('force_schedule_insufficient_hc', False)

        all_solutions = {}
        jobs_to_process = jobs_for_preview

        with st.spinner("Solving schedules and calculating costs... This may take a few minutes."):
            progress_bar = st.progress(0)
            for i, job in enumerate(jobs_to_process):
                display_name = job.get('display')
                if 'Scenario' in job:
                    scenario, week_start_str = job['Scenario'], job['Week_Start_Day']
                    key = f"{scenario} | Week of {week_start_str}"
                    week_start_dt = datetime.datetime.strptime(week_start_str, '%Y-%m-%d').date()
                    scenario_data = st.session_state.all_scenarios[scenario]
                    full_staffing_df, forecast_params = scenario_data
                    week_dates = pd.to_datetime([(week_start_dt + datetime.timedelta(days=d)) for d in range(7)])
                    req_pivot = full_staffing_df.pivot_table(index='Date', columns='Interval', values='final_positions', fill_value=0).reindex(index=week_dates, fill_value=0)
                    base_required_staff_matrix = np.ceil(req_pivot.values).astype(int).tolist()
                else:
                    week_start_dt = job['week_start_dt']
                    key = f"Manual Input | Week of {week_start_dt.strftime('%Y-%m-%d')}"
                    display_name = key
                    base_required_staff_matrix = job['matrix']
                    full_staffing_df, forecast_params = None, None

                num_days_in_week = len(base_required_staff_matrix)
                num_intervals_in_day = len(base_required_staff_matrix[0]) if num_days_in_week > 0 else 0
                inflated_req_matrix = [[0] * num_intervals_in_day for _ in range(num_days_in_week)]
                for d in range(num_days_in_week):
                    current_day_name = days_of_week_ordered[d]
                    day_shrinkage_percents = st.session_state.daily_shrinkage_df[current_day_name].tolist()
                    for p in range(num_intervals_in_day):
                        original_req = base_required_staff_matrix[d][p]
                        shrinkage_percent = day_shrinkage_percents[p]
                        denominator = 1 - (shrinkage_percent / 100.0)
                        inflated_req = math.ceil(original_req / denominator) if (original_req > 0 and denominator > 0) else original_req
                        inflated_req_matrix[d][p] = int(inflated_req)

                requirements_to_store = {'base': base_required_staff_matrix, 'inflated': inflated_req_matrix}

                all_solutions[key] = {}
                if schedule_generation_mode == "Use Pre-defined Shifts":
                    if force_schedule_flag:
                        st.info(f"For job '{key}', scheduling with insufficient headcount is enabled. Only the relaxed 'Best Fit' model will be run.")
                        model_types_to_run = ['best_fit']
                        line_adherence_config = None
                    else:
                        model_types_to_run = ['best_fit', 'meet_or_exceed']
                        if st.session_state.enable_adherence:
                            model_types_to_run.append('line_adherence')
                            line_adherence_config = {'target_level': st.session_state.adherence_target_level.lower(), 'target_percent': st.session_state.adherence_target_percent, 'cap_percent': st.session_state.adherence_cap_percent}
                        else:
                            line_adherence_config = None

                    try:
                        virtual_shifts, shift_groups = expand_shifts_for_solver(st.session_state.shifts_df)
                        if not virtual_shifts:
                            st.error(f"For job '{key}', no valid shifts could be created. Please check your shift definitions in the sidebar."); continue
                    except Exception as e:
                        st.error(f"For job '{key}', error processing shift definitions: {e}"); continue

                    num_employees = st.session_state.adjusted_headcounts.get(display_name, math.ceil(job.get('Required HC (Avg)', job.get('avg_fte', 0))))

                    for model_type in model_types_to_run:
                        solution = solve_schedule_ortools(
                            inflated_req_matrix, virtual_shifts, shift_groups, num_employees, work_days_by_shift,
                            model_type, constraints_config, days_of_week_ordered,
                            line_adherence_config=(line_adherence_config if model_type == 'line_adherence' else None),
                            force_fit_mode=force_schedule_flag
                        )
                        if solution.get('status') in ['OPTIMAL', 'FEASIBLE']:
                            cost_breakdown, cost_details, weekly_hours_df = calculate_schedule_cost(solution['roster_df'], virtual_shifts, cost_config, week_start_dt, days_of_week_ordered)
                            solution['config'] = line_adherence_config if model_type == 'line_adherence' else None
                            all_solutions[key][model_type] = {'solution': solution, 'requirements': requirements_to_store, 'cost': (cost_breakdown, cost_details, weekly_hours_df), 'virtual_shifts_used': virtual_shifts, 'forecast_df': full_staffing_df, 'forecast_params': forecast_params}
                        else:
                            all_solutions[key][model_type] = {'solution': solution, 'requirements': requirements_to_store, 'cost': (None, None, None), 'virtual_shifts_used': virtual_shifts, 'forecast_df': full_staffing_df, 'forecast_params': forecast_params}

                else: # "Optimize Shifts Automatically"
                    optimization_config = {
                        'shift_consistency': st.session_state.get('shift_consistency_opt', False),
                        'max_unique_shifts': st.session_state.max_unique_shifts,
                        'daily_op_hours': st.session_state.daily_op_hours,
                        'allowed_durations': st.session_state.allowed_durations,
                        'duration_rules': st.session_state.duration_rules,
                        'total_headcount': st.session_state.total_hc_optimization,
                        'min_agents_per_shift': st.session_state.min_agents_per_shift,
                        'max_agents_per_shift': st.session_state.max_agents_per_shift,
                        'distribution_caps': st.session_state.distribution_caps,
                        'force_fit': force_schedule_flag # Pass the flag to the optimizer
                    }

                    chosen_model_type = st.session_state.optimization_model_choice
                    peak_req = np.max(np.array(inflated_req_matrix))
                    optimizer_hc = st.session_state.total_hc_optimization
                    model_to_run = 'line_adherence' if chosen_model_type == "Line Adherence (Coverage Target)" else 'best_fit'

                    if model_to_run == 'line_adherence' and optimizer_hc < peak_req:
                        if force_schedule_flag:
                            st.warning(f"For job '{key}', headcount ({optimizer_hc}) is insufficient for peak demand ({int(peak_req)}). Overriding 'Line Adherence' model and running relaxed 'Best Fit' instead to find the best possible partial coverage.", icon="⚠️")
                            model_to_run = 'best_fit'
                        else:
                            st.error(f"Cannot run 'Line Adherence' optimizer for job '{key}': headcount ({optimizer_hc}) is less than peak requirement ({int(peak_req)}). Increase headcount or check 'Attempt to schedule...' in the sidebar to run a 'Best Fit' schedule.", icon="🛑")
                            continue

                    if model_to_run == 'line_adherence':
                        optimization_config['model_type'] = 'line_adherence'
                        optimization_config['target_level'] = st.session_state.opt_adherence_target_level.lower()
                        optimization_config['target_percent'] = st.session_state.opt_adherence_target_percent
                        optimization_config['cap_percent'] = st.session_state.opt_adherence_cap_percent
                        model_key_suffix = "line_adherence"
                    else: # best_fit
                        optimization_config['model_type'] = 'best_fit'
                        optimization_config['understaff_penalty'] = st.session_state.understaff_penalty_opt
                        optimization_config['overstaff_penalty'] = st.session_state.overstaff_penalty_opt
                        model_key_suffix = "best_fit"

                    solution = solve_schedule_with_shift_optimization(inflated_req_matrix, days_of_week_ordered, optimization_config)
                    model_type = f"shift_optimization_{model_key_suffix}"

                    if solution.get('status') in ['OPTIMAL', 'FEASIBLE']:
                        optimized_virtual_shifts = solution['virtual_shifts_generated']
                        cost_breakdown, cost_details, weekly_hours_df = calculate_schedule_cost(solution['roster_df'], optimized_virtual_shifts, cost_config, week_start_dt, days_of_week_ordered)
                        solution['config'] = {
                            'target_level': optimization_config.get('target_level'),
                            'target_percent': optimization_config.get('target_percent'),
                            'cap_percent': optimization_config.get('cap_percent')
                        } if model_key_suffix == "line_adherence" else None

                        all_solutions[key][model_type] = {'solution': solution, 'requirements': requirements_to_store, 'cost': (cost_breakdown, cost_details, weekly_hours_df), 'virtual_shifts_used': optimized_virtual_shifts, 'forecast_df': full_staffing_df, 'forecast_params': forecast_params}
                    else:
                        all_solutions[key][model_type] = {'solution': solution, 'requirements': requirements_to_store, 'cost': (None, None, None), 'virtual_shifts_used': [], 'forecast_df': full_staffing_df, 'forecast_params': forecast_params}
                progress_bar.progress((i + 1) / len(jobs_to_process))

        st.session_state.scheduling_solutions = {'solutions': all_solutions}
        end_time = time.time()
        st.success(f"Finished generating and costing schedules! Time taken: {format_duration(time.time() - start_time)}")


    if 'scheduling_solutions' in st.session_state and st.session_state.scheduling_solutions:
        st.header("Generated Schedule Results")
        solutions_data = st.session_state.scheduling_solutions.get('solutions', {})
        for key, solutions in solutions_data.items():
            with st.expander(f"**Results for: {key}**", expanded=True):
                tabs_to_create = []

                model_titles = {
                    'shift_optimization_best_fit': "⭐ Shift Optimization (Best Fit)",
                    'shift_optimization_line_adherence': "⭐ Shift Optimization (Line Adherence)",
                    'best_fit': "Best Fit Model (Balanced)",
                    'meet_or_exceed': "Meet or Exceed Model (Coverage-Focused)",
                    'line_adherence': f"Line Adherence ({solutions.get('line_adherence',{}).get('solution',{}).get('config',{}).get('target_percent','N/A')}%, {solutions.get('line_adherence',{}).get('solution',{}).get('config',{}).get('target_level','').capitalize()})"
                }

                for model_key, title in model_titles.items():
                    if model_key in solutions: # Check if model was run
                        tabs_to_create.append((model_key, title))

                if not tabs_to_create:
                     st.error("No models were run for this job. Please check your settings.")
                     continue

                tab_names = [title for _, title in tabs_to_create]
                if not any(sol.get('solution', {}).get('status') in ['OPTIMAL', 'FEASIBLE'] for sol in solutions.values()):
                    st.error(f"Solver did not find a feasible solution for any model for this job. This usually means the demand is impossible to meet with the current staff pool and constraints. Try increasing the number of employees or relaxing constraints.")
                    for model_key, data in solutions.items():
                        st.write(f"**{model_titles.get(model_key, model_key).split(' (')[0]} Status:** {data.get('solution',{}).get('status', 'NOT RUN')}")
                        reason = data.get('solution', {}).get('reason')
                        if reason:
                            st.write(f"Reason: {reason}")
                    continue

                successful_tabs = [(mk, mt) for mk, mt in tabs_to_create if solutions[mk].get('solution', {}).get('status') in ['OPTIMAL', 'FEASIBLE']]

                if not successful_tabs:
                    st.error("No model found a successful solution.")
                    continue

                tab_names = [title for _, title in successful_tabs]
                tabs = st.tabs(tab_names)

                for i, (model_key, title) in enumerate(successful_tabs):
                    with tabs[i]:
                        solution_data_full = solutions.get(model_key)

                        if model_key == 'shift_optimization_best_fit':
                            subheader_text = "Objective: Find the best shifts and roster to minimize weighted over/understaffing."
                        elif model_key == 'shift_optimization_line_adherence':
                            config = solution_data_full.get('solution', {}).get('config', {})
                            subheader_text = f"Objective: Meet minimum {config.get('target_percent', 'N/A')}% adherence at the {config.get('target_level', '').capitalize()} level with minimum staff."
                        elif model_key == 'best_fit':
                            subheader_text = "Objective: Minimize Weighted Over/Understaffing"
                        elif model_key == 'meet_or_exceed':
                            subheader_text = "Objective: Fulfill All Requirements (100% Coverage), Minimize Surplus"
                        elif model_key == 'line_adherence':
                            config = solution_data_full.get('solution', {}).get('config', {})
                            subheader_text = f"Objective: Meet minimum {config.get('target_percent', 'N/A')}% adherence at the {config.get('target_level', '').capitalize()} level with minimum cost."

                        st.subheader(subheader_text)

                        display_comprehensive_results(
                            solution_data_full['solution'],
                            solution_data_full['cost'],
                            solution_data_full['requirements'],
                            f"{model_key}_{sanitize_name(key)}",
                            days_of_week_ordered,
                            solution_data_full['virtual_shifts_used']
                        )


with tab3:
    st.header("Results & Performance Summary")
    st.subheader("Weekly Staffing Requirement Summary")
    if "scenario_summary" in st.session_state and not st.session_state.scenario_summary.empty:
        st.dataframe(st.session_state.scenario_summary.style.format({
            "Required HC (Avg)": '{:.2f}',
            "Required HC for Peak Day": '{:.2f}',
            "Total Volume": '{:,.0f}',
            "Weekly ASA (s)": '{:.2f}',
            "Weekly SL (%)": '{:.2f}%',
            "Weekly Occ. (%)": '{:.2f}%',
            "Weekly Wait Prob. (%)": '{:.2f}%',
        }, na_rep="N/A"), use_container_width=True)
        download_dataframe_csv(st.session_state.scenario_summary, "summary_tab3_scenario_summary")
    else: st.info("Run the Staffing Calculator on Tab 1 to see this.")

    st.subheader("Scheduling Model Performance & Cost Summary")
    summary_data = []
    if "scheduling_solutions" in st.session_state and st.session_state.scheduling_solutions:
        solutions_data = st.session_state.scheduling_solutions.get('solutions', {})
        interval_duration_hours = pd.to_timedelta(st.session_state.interval_freq).total_seconds() / 3600
        for key, solutions in solutions_data.items():
            for model_type, data in solutions.items():
                model_name_map = {
                    'best_fit': 'Best Fit',
                    'meet_or_exceed': 'Meet or Exceed',
                    'shift_optimization_best_fit': '⭐ Opti (Best Fit)',
                    'shift_optimization_line_adherence': '⭐ Opti (Line Adherence)'
                }
                model_name = model_name_map.get(model_type, model_type.replace('_', ' ').title())

                config = data.get('solution', {}).get('config', {})
                if model_type == 'line_adherence' and config:
                    model_name = f"Line Adherence ({config.get('target_percent')}%, {config.get('target_level').capitalize()})"

                if " | Week of " in key:
                    scenario_name, week_start_str = key.split(" | Week of ", 1)
                else:
                    scenario_name = key
                    week_start_str = "N/A"

                row = {"Scenario": scenario_name, "Week Starting": week_start_str, "Model Type": model_name}

                if (data and data.get('solution', {}).get('status') in ['OPTIMAL', 'FEASIBLE'] and 'cost' in data and data['cost'][0] is not None):

                    req_data = data['requirements']
                    if isinstance(req_data, dict) and 'inflated' in req_data:
                        inflated_req_matrix = req_data['inflated']
                    else:
                        inflated_req_matrix = req_data

                    fte_metrics = calculate_fte_metrics_from_matrix(
                        inflated_req_matrix,
                        st.session_state.get('working_hours', 8.0),
                        st.session_state.get('working_days', 5.0)
                    )
                    row.update({
                        "Inflated HC (Avg)": fte_metrics['avg_fte'],
                        "Inflated HC (Peak)": fte_metrics['peak_fte']
                    })

                    req = np.array(inflated_req_matrix)
                    sched = np.array(data['solution']['scheduled_staff'])
                    diff, total_req_intervals = sched - req, np.sum(req)
                    over, under = np.sum(diff[diff > 0]), -np.sum(diff[diff < 0])
                    coverage = (np.sum(sched) - over) / total_req_intervals if total_req_intervals > 0 else 1.0
                    cost_breakdown, _, _ = data['cost']

                    total_cost = cost_breakdown['Total']
                    total_scheduled_intervals = np.sum(sched)
                    total_scheduled_hours = total_scheduled_intervals * interval_duration_hours
                    total_required_hours = total_req_intervals * interval_duration_hours

                    original_staffing_df = data.get('forecast_df')
                    if original_staffing_df is not None and not original_staffing_df.empty:
                        week_start_dt_for_volume = pd.to_datetime(week_start_str)
                        week_df_for_volume = original_staffing_df[original_staffing_df['Week_Start_Day'] == week_start_dt_for_volume]
                        total_volume = week_df_for_volume['Volume'].sum()
                        row["Cost/Transaction"] = total_cost / total_volume if total_volume > 0 else 0
                    else:
                        row["Cost/Transaction"] = np.nan


                    row["Cost/Sched Hour"] = total_cost / total_scheduled_hours if total_scheduled_hours > 0 else 0
                    row["Cost/Req Hour"] = total_cost / total_required_hours if total_required_hours > 0 else 0

                    cost_breakdown_no_total = cost_breakdown.copy()
                    cost_breakdown_no_total.pop('Total', None)
                    base_cost = cost_breakdown_no_total.get('Base Pay', 0)
                    ot_cost = cost_breakdown_no_total.get('Overtime Premium', 0)
                    diff_cost = cost_breakdown_no_total.get('Shift Differentials', 0) + cost_breakdown_no_total.get('Day/Holiday Premiums', 0)

                    row.update({"Status": data['solution']['status'], "Total Cost": total_cost,
                                "Coverage Met": f"{coverage:.2%}", "Understaffed Intervals": under, "Overstaffed Intervals": over,
                                "VTO Hours (Opportunities)": over * interval_duration_hours, "OT Hours (Needed)": under * interval_duration_hours,
                                "Base Cost": base_cost, "OT Cost": ot_cost, "Differential Cost": diff_cost})
                else:
                    status = data.get('solution', {}).get('status', 'Failed') if data else 'Not Run'
                    row.update({"Status": status, "Total Cost": np.nan, "Coverage Met": np.nan,
                                "Understaffed Intervals": np.nan, "Overstaffed Intervals": np.nan,
                                "VTO Hours (Opportunities)": np.nan, "OT Hours (Needed)": np.nan,
                                "Base Cost": np.nan, "OT Cost": np.nan, "Differential Cost": np.nan,
                                "Inflated HC (Avg)": np.nan, "Inflated HC (Peak)": np.nan,
                                "Cost/Sched Hour": np.nan, "Cost/Req Hour": np.nan, "Cost/Transaction": np.nan})
                summary_data.append(row)

        if summary_data:
            summary_display_df = pd.DataFrame(summary_data)
            cols_order = [
                "Scenario", "Week Starting", "Model Type", "Status", "Total Cost",
                "Cost/Sched Hour", "Cost/Req Hour", "Cost/Transaction",
                "Inflated HC (Avg)", "Inflated HC (Peak)", "Coverage Met",
                "Understaffed Intervals", "Overstaffed Intervals",
                "VTO Hours (Opportunities)", "OT Hours (Needed)",
                "Base Cost", "OT Cost", "Differential Cost"
                ]

            final_cols = [col for col in cols_order if col in summary_display_df.columns]
            summary_display_df = summary_display_df[final_cols]

            st.dataframe(summary_display_df.style.format({
                'Total Cost': '${:,.2f}', 'Base Cost': '${:,.2f}', 'OT Cost': '${:,.2f}',
                'Differential Cost': '${:,.2f}', 'Understaffed Intervals': '{:,.0f}',
                'Overstaffed Intervals': '{:,.0f}',
                'VTO Hours (Opportunities)': '{:,.1f}', 'OT Hours (Needed)': '{:,.1f}',
                'Inflated HC (Avg)': '{:.2f}', 'Inflated HC (Peak)': '{:.2f}',
                'Cost/Sched Hour': '${:,.2f}', 'Cost/Req Hour': '${:,.2f}', 'Cost/Transaction': '${:,.2f}',
                'Coverage Met': '{}'
            }, na_rep='-'), use_container_width=True)
            download_dataframe_csv_no_index(summary_display_df, "scheduling_performance_summary")

    else:
        st.info("Generate weekly rosters on Tab 2 to see a summary here.")

    st.markdown("---")

    st.markdown("---")
    st.header("Head-to-Head Scenario Comparison")

    if 'scheduling_solutions' in st.session_state and st.session_state.scheduling_solutions:
        solutions_data = st.session_state.scheduling_solutions.get('solutions', {})
        valid_solutions = []
        for key, solutions in solutions_data.items():
            for model_type, data in solutions.items():
                if (data and data.get('solution', {}).get('status') in ['OPTIMAL', 'FEASIBLE'] and 'cost' in data and data['cost'][0] is not None):
                    model_name_map = {
                        'best_fit': 'Best Fit',
                        'meet_or_exceed': 'Meet or Exceed',
                        'shift_optimization_best_fit': '⭐ Opti (Best Fit)',
                        'shift_optimization_line_adherence': '⭐ Opti (Line Adherence)'
                    }
                    model_name = model_name_map.get(model_type, model_type.replace('_', ' ').title())
                    config = data.get('solution', {}).get('config', {})
                    if model_type == 'line_adherence' and config:
                        model_name = f"Line Adherence ({config.get('target_percent')}%, {config.get('target_level').capitalize()})"
                    display_name = f"{key} - {model_name}"
                    valid_solutions.append({'display': display_name, 'key': key, 'model': model_type})

        if len(valid_solutions) < 2:
            st.info("You need at least two successful schedule generations to use the comparison tool.")
        else:
            options = [s['display'] for s in valid_solutions]

            def get_solution_data(selection_display_name):
                for sol in valid_solutions:
                    if sol['display'] == selection_display_name:
                        return st.session_state.scheduling_solutions['solutions'][sol['key']][sol['model']]
                return None

            def display_summary_card(column, data, day_order, key_suffix, selection_name):
                with column:
                    st.subheader(selection_name.split(' - ')[0])
                    st.caption(f"Model: {selection_name.split(' - ')[1]}")

                    if not data:
                        st.warning("Could not load data for this scenario."); return None

                    req_data = data['requirements']
                    req_matrix = np.array(req_data['inflated']) if isinstance(req_data, dict) and 'inflated' in req_data else np.array(req_data)
                    sched_matrix = np.array(data['solution']['scheduled_staff'])

                    interval_duration_hours = pd.to_timedelta(st.session_state.interval_freq).total_seconds() / 3600
                    diff_matrix, total_req = sched_matrix - req_matrix, np.sum(req_matrix)
                    over = np.sum(diff_matrix[diff_matrix > 0])
                    under = -np.sum(diff_matrix[diff_matrix < 0])
                    coverage = (np.sum(sched_matrix) - over) / total_req if total_req > 0 else 1.0
                    cost_breakdown, _, _ = data['cost']
                    vto_hours = over * interval_duration_hours
                    ot_needed_hours = under * interval_duration_hours

                    config_dict = (data.get('solution') or {}).get('config') or {}
                    adherence_cap = config_dict.get('cap_percent', 105)
                    
                    adherence_metrics = calculate_adherence_metrics(req_matrix, sched_matrix, adherence_cap, day_order)

                    kpi_data = {
                        'Metric': ['Total Labor Cost', 'Coverage Met', f"Weekly Adherence (Capped @{adherence_cap}%)", 'VTO Hours', 'OT Needed Hours'],
                        'Value': [cost_breakdown['Total'], coverage, adherence_metrics['weekly_adherence'], vto_hours, ot_needed_hours]
                    }
                    summary_kpi_df = pd.DataFrame(kpi_data).set_index("Metric")
                    st.dataframe(summary_kpi_df.style.format({
                        'Value': lambda x: f"${x:,.2f}" if "Cost" in summary_kpi_df.loc[summary_kpi_df['Value'] == x].index[0] else (f"{x:.2%}" if "Coverage" in summary_kpi_df.loc[summary_kpi_df['Value'] == x].index[0] else (f"{x:.2f}%" if "Adherence" in summary_kpi_df.loc[summary_kpi_df['Value'] == x].index[0] else f"{x:,.1f}"))
                    }), use_container_width=True)

                    st.markdown("###### Cost Breakdown")
                    st.dataframe(pd.DataFrame.from_dict(cost_breakdown, orient='index', columns=['Amount']).style.format('${:,.2f}'))

                    st.markdown("##### Over/Understaffing Heatmap")
                    fig_diff = go.Figure(data=go.Heatmap(
                        z=diff_matrix.T, x=day_order, y=[t.strftime('%H:%M') for t in st.session_state.intervals],
                        colorscale='RdBu', zmid=0))
                    fig_diff.update_yaxes(autorange='reversed')
                    fig_diff.update_layout(title="Scheduled vs. Inflated Requirement", title_x=0.5, height=300, margin=dict(l=20, r=20, t=40, b=20))
                    st.plotly_chart(fig_diff, use_container_width=True, key=f"compare_heatmap_{key_suffix}")
                    return summary_kpi_df

            with st.columns(2)[0]:
                selection_A = st.selectbox("Compare Scenario A", options, index=0, key="compare_A")
            with st.columns(2)[1]:
                selection_B = st.selectbox("Compare Scenario B", options, index=1 if len(options)>1 else 0, key="compare_B")

            if selection_A and selection_B:
                if selection_A == selection_B:
                    st.warning("Please select two different scenarios to compare.")
                else:
                    data_A, data_B = get_solution_data(selection_A), get_solution_data(selection_B)

                    if data_A and data_B:
                        disp_col1, disp_col2 = st.columns(2)
                        summary_kpi_df_A = display_summary_card(disp_col1, data_A, days_of_week_ordered, "A", selection_A)
                        summary_kpi_df_B = display_summary_card(disp_col2, data_B, days_of_week_ordered, "B", selection_B)

                        st.markdown("---")
                        st.subheader("Key Metric Comparison Charts")

                        metrics_for_charting = [
                            ('Total Labor Cost', 'Total Labor Cost'),
                            ('Coverage Met', 'Coverage Met'),
                            ('VTO Hours', 'VTO Hours'),
                            ('OT Needed Hours', 'OT Needed Hours')
                        ]

                        chart_data_rows = []
                        for metric_key, display_label in metrics_for_charting:
                            val_A_num = summary_kpi_df_A.loc[display_label, 'Value'] if summary_kpi_df_A is not None and display_label in summary_kpi_df_A.index else 0
                            val_B_num = summary_kpi_df_B.loc[display_label, 'Value'] if summary_kpi_df_B is not None and display_label in summary_kpi_df_B.index else 0
                            chart_data_rows.append({'Metric': metric_key, selection_A: val_A_num, selection_B: val_B_num})

                        comparison_df_for_plot = pd.DataFrame(chart_data_rows).set_index('Metric')

                        fig_compare_metrics = go.Figure(data=[
                            go.Bar(name=selection_A, x=comparison_df_for_plot.index, y=comparison_df_for_plot[selection_A]),
                            go.Bar(name=selection_B, x=comparison_df_for_plot.index, y=comparison_df_for_plot[selection_B])
                        ])
                        fig_compare_metrics.update_layout(
                            barmode='group',
                            title='Key Performance Indicators Comparison',
                            yaxis_title='Value',
                            legend_title='Scenario'
                        )
                        st.plotly_chart(fig_compare_metrics, use_container_width=True)
                        download_dataframe_csv(comparison_df_for_plot, "key_metric_comparison")

                    else:
                        st.error("Could not load complete data for one or both selected solutions.")
    else:
        st.info("Generate schedules in Tab 2 to enable the comparison tool.")

with tab4:
    st.header("Forecast 'What-If' Simulation")
    st.info(
        "Instantly visualize how operational changes would impact your staffing requirements. "
        "Select a baseline forecast from Tab 1, then use the sliders to simulate real-time events "
        "like volume spikes, AHT changes, or shrinkage adjustments."
    )
    st.markdown("---")

    forecast_data = st.session_state.get('all_scenarios', {})
    if not forecast_data:
        st.warning("No forecast scenarios found. Please run a staffing calculation on Tab 1 to enable this feature.", icon="⚠️")
    else:
        left_col, right_col = st.columns(2)

        with left_col:
            st.subheader("Simulation Controls")
            # Create a list of available forecasts to choose from
            forecast_options = list(forecast_data.keys())
            selected_scenario_name = st.selectbox(
                "Select a baseline forecast scenario to simulate:",
                options=forecast_options,
                index=0,
                key="what_if_forecast_selection"
            )

            # Get the data for the selected scenario
            baseline_staffing_df, baseline_params = forecast_data[selected_scenario_name]
            is_blended = 'channels' in baseline_params

            # Allow user to select a specific week from the chosen scenario
            week_start_options = sorted(baseline_staffing_df['Week_Start_Day'].dt.strftime('%Y-%m-%d').unique())
            if not week_start_options:
                 st.warning("No weekly data available for the selected scenario.")
                 st.stop()

            selected_week_str = st.selectbox(
                "Select a week to simulate:",
                options=week_start_options,
                key="what_if_week_selection"
            )
            
            volume_percent = st.slider("Volume Fluctuation (%)", -100, 100, 0, 5, key="what_if_volume", help="Increases or decreases the original contact volume forecast.")
            aht_percent = st.slider("AHT Fluctuation (%)", -50, 50, 0, 5, key="what_if_aht", help="Increases or decreases the original AHT forecast.")
            shrink_adj = st.slider("Shrinkage Adjustment (Absolute %)", -20.0, 20.0, 0.0, 0.5, key="what_if_shrink", help="Adds or subtracts from the original shrinkage percentage. E.g., if original is 30% and slider is +5%, new shrinkage is 35%.")
            week_start_dt = datetime.datetime.strptime(selected_week_str, '%Y-%m-%d').date()

        with right_col:
            st.subheader("Simulation Results")
            # --- Recalculate the forecast based on slider inputs ---
            with st.spinner("Recalculating forecast based on new inputs..."):
                simulated_params = copy.deepcopy(baseline_params)
                volume_multiplier = 1 + (volume_percent / 100.0)
                if not is_blended:
                    simulated_params['volume_adjustment'] = simulated_params.get('volume_adjustment', 100) * volume_multiplier
                else:
                    pass
                aht_multiplier = 1 + (aht_percent / 100.0)
                if not is_blended:
                    simulated_params['aht'] = simulated_params['aht'] * aht_multiplier
                else:
                    for ch_name in simulated_params['channel_params']:
                        simulated_params['channel_params'][ch_name]['aht'] = simulated_params['channel_params'][ch_name]['aht'] * aht_multiplier
                if not is_blended:
                    original_shrinkage = simulated_params.get('shrinkage', 0)
                    simulated_params['shrinkage'] = max(0, min(99.9, original_shrinkage + shrink_adj))
                else:
                    for ch_name in simulated_params['channel_params']:
                        original_shrinkage = simulated_params['channel_params'][ch_name].get('shrinkage', 0)
                        simulated_params['channel_params'][ch_name]['shrinkage'] = max(0, min(99.9, original_shrinkage + shrink_adj))
                week_mask = (baseline_staffing_df['Week_Start_Day'] == pd.Timestamp(week_start_dt))
                baseline_week_df = baseline_staffing_df[week_mask].copy()

                if not is_blended:
                    week_dates = baseline_week_df['Date'].dt.strftime('%Y-%m-%d').unique()
                    day_name_map_sim = {date: pd.to_datetime(date).strftime('%A') for date in week_dates}
                    volume_data = {}
                    for date in week_dates:
                        date_obj = pd.to_datetime(date)
                        volume_data[date] = baseline_week_df[baseline_week_df['Date'] == date_obj].set_index('Interval')['Volume'] * volume_multiplier
                    volume_df_sim = pd.DataFrame(volume_data)
                    volume_df_sim.index = baseline_week_df[baseline_week_df['Date'] == pd.to_datetime(week_dates[0])]['Interval'].values
                    simulated_staffing_df = run_staffing_calculation(simulated_params, week_dates, day_name_map_sim, st.session_state.week_start_day, volume_df_sim)
                else:
                    simulated_staffing_df = baseline_week_df.copy()
                    simulated_staffing_df['Volume'] = simulated_staffing_df['Volume'] * volume_multiplier
                    simulated_staffing_df['AHT'] = simulated_staffing_df['AHT'] * aht_multiplier
                    simulated_staffing_df['Simulated Shrinkage Adjustment'] = shrink_adj
                if not is_blended:
                    req_pivot_sim = simulated_staffing_df.pivot_table(index='Interval', columns='Date', values='final_positions', fill_value=0)
                else:
                    req_pivot_sim = simulated_staffing_df.pivot_table(index='Interval', columns='Date', values='final_positions', fill_value=0)
                req_pivot_sim = req_pivot_sim.reindex(index=st.session_state.intervals, fill_value=0)
                week_dates_dt = pd.to_datetime(week_dates) if not is_blended else simulated_staffing_df['Date'].unique()
                req_pivot_sim = req_pivot_sim.reindex(columns=sorted(week_dates_dt), fill_value=0)
                days_of_week_ordered_sim = [pd.to_datetime(date).strftime('%A') for date in sorted(week_dates)]
                if not is_blended:
                    volume_pivot_sim = simulated_staffing_df.pivot_table(index='Interval', columns='Date', values='Volume', fill_value=0)
                    volume_pivot_sim = volume_pivot_sim.reindex(index=st.session_state.intervals, fill_value=0)
                    volume_pivot_sim = volume_pivot_sim.reindex(columns=sorted(week_dates_dt), fill_value=0)
                else:
                    volume_pivot_sim = simulated_staffing_df.pivot_table(index='Interval', columns='Date', values='Volume', fill_value=0)
                    volume_pivot_sim = volume_pivot_sim.reindex(index=st.session_state.intervals, fill_value=0)
                    volume_pivot_sim = volume_pivot_sim.reindex(columns=sorted(simulated_staffing_df['Date'].unique()), fill_value=0)

            # Get baseline data for comparison
            baseline_week_mask = (baseline_staffing_df['Week_Start_Day'] == pd.Timestamp(week_start_dt))
            baseline_week_df_comp = baseline_staffing_df[baseline_week_mask].copy()
            req_pivot_baseline = baseline_week_df_comp.pivot_table(index='Interval', columns='Date', values='final_positions', fill_value=0)
            req_pivot_baseline = req_pivot_baseline.reindex(index=st.session_state.intervals, fill_value=0)
            req_pivot_baseline = req_pivot_baseline.reindex(columns=sorted(week_dates_dt), fill_value=0)
            volume_pivot_baseline = baseline_week_df_comp.pivot_table(index='Interval', columns='Date', values='Volume', fill_value=0)
            volume_pivot_baseline = volume_pivot_baseline.reindex(index=st.session_state.intervals, fill_value=0)
            volume_pivot_baseline = volume_pivot_baseline.reindex(columns=sorted(week_dates_dt), fill_value=0)
            req_diff = req_pivot_sim - req_pivot_baseline
            volume_diff = volume_pivot_sim - volume_pivot_baseline
            intervals_str_fmt = [t.strftime('%H:%M') for t in st.session_state.intervals]
            
            # Display summary metrics
            col1, col2, col3, col4 = st.columns(4)
            total_baseline_req = req_pivot_baseline.sum().sum()
            total_simulated_req = req_pivot_sim.sum().sum()
            req_change = ((total_simulated_req - total_baseline_req) / total_baseline_req) * 100 if total_baseline_req > 0 else 0
            total_baseline_vol = volume_pivot_baseline.sum().sum()
            total_simulated_vol = volume_pivot_sim.sum().sum()
            vol_change = ((total_simulated_vol - total_baseline_vol) / total_baseline_vol) * 100 if total_baseline_vol > 0 else 0
            col1.metric("Total Required Staff", f"{total_simulated_req:,.0f}", f"{req_change:.1f}%")
            col2.metric("Total Volume", f"{total_simulated_vol:,.0f}", f"{vol_change:.1f}%")
            if not is_blended:
                baseline_aht = baseline_params.get('aht', 0)
                simulated_aht = simulated_params.get('aht', 0)
                aht_change = ((simulated_aht - baseline_aht) / baseline_aht) * 100 if baseline_aht > 0 else 0
                col3.metric("AHT", f"{simulated_aht:.1f}s", f"{aht_change:.1f}%")
            else:
                col3.metric("AHT Change", f"{aht_percent}%", "Multi-channel")
            if not is_blended:
                baseline_shrinkage = baseline_params.get('shrinkage', 0)
                simulated_shrinkage = simulated_params.get('shrinkage', 0)
                shrinkage_change = simulated_shrinkage - baseline_shrinkage
                col4.metric("Shrinkage", f"{simulated_shrinkage:.1f}%", f"{shrinkage_change:+.1f}%")
            else:
                col4.metric("Shrinkage Change", f"{shrink_adj:+.1f}%", "Multi-channel")
            
            # --- Organize outputs into tabs ---
            tab_summary, tab_charts, tab_params, tab_interval_tables, tab_heatmaps = st.tabs([
                "📊 Summary KPIs", "📈 Comparison Charts", "📋 Parameter Details", "🗓️ Interval Staffing Tables", "🔥 Difference Heatmaps"
            ])

            with tab_summary:
                st.subheader("Key Performance Indicator Summary")
                st.info("This table compares the key headcount metrics between the baseline forecast and your simulated scenario for the selected week.")

                working_hours = st.session_state.get('working_hours', 8.0)
                working_days = st.session_state.get('working_days', 5.0)
                interval_duration_hours = pd.to_timedelta(st.session_state.interval_freq).total_seconds() / 3600

                # --- Calculate Baseline Metrics ---
                baseline_fte_metrics = calculate_fte_metrics_from_matrix(req_pivot_baseline.T.values.tolist(), working_hours, working_days)
                baseline_daily_hours = req_pivot_baseline.sum(axis=0) * interval_duration_hours
                baseline_daily_hc = baseline_daily_hours / working_hours if working_hours > 0 else pd.Series(0.0, index=baseline_daily_hours.index)

                # --- Calculate Simulated Metrics ---
                simulated_fte_metrics = calculate_fte_metrics_from_matrix(req_pivot_sim.T.values.tolist(), working_hours, working_days)
                simulated_daily_hours = req_pivot_sim.sum(axis=0) * interval_duration_hours
                simulated_daily_hc = simulated_daily_hours / working_hours if working_hours > 0 else pd.Series(0.0, index=simulated_daily_hours.index)

                # --- Construct Summary DataFrame ---
                summary_rows = []
                summary_rows.append({
                    "Metric": "Weekly Required HC (Avg)",
                    "Baseline": baseline_fte_metrics['avg_fte'],
                    "Simulated": simulated_fte_metrics['avg_fte']
                })
                summary_rows.append({
                    "Metric": "Required HC for Peak Day",
                    "Baseline": baseline_fte_metrics['peak_fte'],
                    "Simulated": simulated_fte_metrics['peak_fte']
                })

                # Add daily HC, ensuring correct day names
                for day_date in sorted(baseline_daily_hc.index):
                    day_name = day_date.strftime('%A')
                    summary_rows.append({
                        "Metric": f"Daily Required HC ({day_name})",
                        "Baseline": baseline_daily_hc.get(day_date, 0),
                        "Simulated": simulated_daily_hc.get(day_date, 0)
                    })

                # Separate weekly and daily metrics
                weekly_summary_rows = []
                daily_summary_rows = []

                for row in summary_rows:
                    if "Weekly" in row["Metric"] or "Peak Day" in row["Metric"]:
                        weekly_summary_rows.append(row)
                    else:
                        daily_summary_rows.append(row)

                weekly_summary_df = pd.DataFrame(weekly_summary_rows).set_index("Metric")
                weekly_summary_df['Difference'] = weekly_summary_df['Simulated'] - weekly_summary_df['Baseline']
                weekly_summary_df['Difference (%)'] = (weekly_summary_df['Difference'] / weekly_summary_df['Baseline']).replace([np.inf, -np.inf], 0).fillna(0) * 100

                daily_summary_df = pd.DataFrame(daily_summary_rows).set_index("Metric")
                daily_summary_df['Difference'] = daily_summary_df['Simulated'] - daily_summary_df['Baseline']
                daily_summary_df['Difference (%)'] = (daily_summary_df['Difference'] / daily_summary_df['Baseline']).replace([np.inf, -np.inf], 0).fillna(0) * 100

                st.markdown("##### Weekly Required HC")
                st.dataframe(weekly_summary_df.style.format({
                    'Baseline': '{:.2f}',
                    'Simulated': '{:.2f}',
                    'Difference': '{:+.2f}',
                    'Difference (%)': '{:+.2f}%'
                }), use_container_width=True)
                download_dataframe_csv(weekly_summary_df, f"whatif_weekly_hc_summary_{selected_scenario_name}_{selected_week_str}")

                st.markdown("##### Daily Required HC")
                st.dataframe(daily_summary_df.style.format({
                    'Baseline': '{:.2f}',
                    'Simulated': '{:.2f}',
                    'Difference': '{:+.2f}',
                    'Difference (%)': '{:+.2f}%'
                }), use_container_width=True)
                download_dataframe_csv(daily_summary_df, f"whatif_daily_hc_summary_{selected_scenario_name}_{selected_week_str}")


            with tab_charts:
                st.markdown("##### Staffing Requirement Comparison")
                fig_comp_req = make_subplots(rows=7, cols=1, shared_xaxes=True, vertical_spacing=0.03, subplot_titles=days_of_week_ordered)
                for i, day_name in enumerate(days_of_week_ordered):
                    day_date_for_plotting = sorted(week_dates_dt)[i]
                    # Ensure the column exists before trying to access it
                    if pd.Timestamp(day_date_for_plotting) in req_pivot_baseline.columns and pd.Timestamp(day_date_for_plotting) in req_pivot_sim.columns:
                        fig_comp_req.add_trace(go.Scatter(x=intervals_str_fmt, y=req_pivot_baseline[pd.Timestamp(day_date_for_plotting)], mode='lines', name='Baseline Required', line=dict(color='blue', dash='dot')), row=i+1, col=1)
                        fig_comp_req.add_trace(go.Scatter(x=intervals_str_fmt, y=req_pivot_sim[pd.Timestamp(day_date_for_plotting)], mode='lines', name='Simulated Required', line=dict(color='red')), row=i+1, col=1)
                fig_comp_req.update_layout(height=1400, showlegend=True, title_text="Daily Required Staff: Baseline vs Simulated")
                st.plotly_chart(fig_comp_req, use_container_width=True)

                st.markdown("##### Volume Comparison")
                fig_comp_vol = make_subplots(rows=7, cols=1, shared_xaxes=True, vertical_spacing=0.03, subplot_titles=days_of_week_ordered)
                for i, day in enumerate(days_of_week_ordered):
                    day_date = sorted(week_dates_dt)[i]
                    date_col_pd = pd.Timestamp(day_date)
                    if date_col_pd in volume_pivot_baseline.columns and date_col_pd in volume_pivot_sim.columns:
                        fig_comp_vol.add_trace(go.Bar(x=intervals_str_fmt, y=volume_pivot_baseline[date_col_pd], name='Baseline Volume', marker_color='lightblue'), row=i+1, col=1)
                        fig_comp_vol.add_trace(go.Bar(x=intervals_str_fmt, y=volume_pivot_sim[date_col_pd], name='Simulated Volume', marker_color='lightcoral'), row=i+1, col=1)
                fig_comp_vol.update_layout(height=1400, showlegend=True, title_text="Daily Volume: Baseline vs Simulated")
                st.plotly_chart(fig_comp_vol, use_container_width=True)


            with tab_params:
                st.markdown("##### Parameter Adjustments")
                # Check if baseline_params has the correct structure for blending or single channel
                if 'channel_params' in baseline_params and is_blended:
                    # For blended scenarios, show individual channel adjustments
                    param_data = []

                    # Volume Multiplier is global for blended scenarios
                    param_data.append({
                        "Parameter": "Overall Volume Multiplier",
                        "Baseline Value": 1.0,
                        "Simulated Value": volume_multiplier,
                        "Change": f"{volume_multiplier - 1.0:+.2f} ({(volume_multiplier - 1.0)/1.0*100:+.1f}%)"
                    })
                    
                    # AHT and Shrinkage for each channel
                    for ch_name in baseline_params['channels']:
                        base_aht = baseline_params['channel_params'][ch_name].get('aht', 0)
                        sim_aht = simulated_params['channel_params'][ch_name].get('aht', 0)
                        aht_change_val = sim_aht / base_aht if base_aht != 0 else (0 if sim_aht == 0 else float('inf'))
                        aht_percent_change = ((sim_aht - base_aht) / base_aht) * 100 if base_aht != 0 else (0 if sim_aht == 0 else float('inf'))
                        param_data.append({
                            "Parameter": f"{ch_name} AHT Multiplier",
                            "Baseline Value": 1.0, # Multiplier not actual AHT value here
                            "Simulated Value": aht_multiplier,
                            "Change": f"{aht_multiplier - 1.0:+.2f} ({(aht_multiplier - 1.0)/1.0*100:+.1f}%)"
                        })
                        
                        base_shrinkage = baseline_params['channel_params'][ch_name].get('shrinkage', 0)
                        sim_shrinkage = simulated_params['channel_params'][ch_name].get('shrinkage', 0)
                        param_data.append({
                            "Parameter": f"{ch_name} Shrinkage Adjustment",
                            "Baseline Value": 0.0, # Absolute adjustment
                            "Simulated Value": shrink_adj,
                            "Change": f"{sim_shrinkage - base_shrinkage:+.1f} pp"
                        })
                else:
                    # Original logic for single channel
                    param_data = {
                        "Parameter": ["Volume Multiplier", "AHT Multiplier", "Shrinkage Adjustment"],
                        "Baseline Value": [1.0, 1.0, 0.0],
                        "Simulated Value": [volume_multiplier, aht_multiplier, shrink_adj]
                    }

                    # Calculate Change column
                    param_data["Change"] = []
                    for i in range(len(param_data["Parameter"])):
                        baseline_val = param_data["Baseline Value"][i]
                        simulated_val = param_data["Simulated Value"][i]
                        change_val = simulated_val - baseline_val
                        
                        if param_data["Parameter"][i] == "Shrinkage Adjustment":
                            param_data["Change"].append(f"{change_val:+.1f} pp")
                        elif baseline_val != 0:
                            percent_change = (change_val / baseline_val) * 100
                            param_data["Change"].append(f"{change_val:+.2f} ({percent_change:+.1f}%)")
                        else:
                            param_data["Change"].append(f"{change_val:+.2f}") # Handle 0 baseline gracefully

                param_df = pd.DataFrame(param_data)
                
                st.dataframe(param_df.style.format({
                    "Baseline Value": '{:,.2f}',
                    "Simulated Value": '{:,.2f}'
                }, na_rep="N/A"))

            with tab_interval_tables:
                st.subheader("Baseline Required Staffing")
                st.info("The original staffing requirements for each interval from the selected forecast.")
                st.dataframe(req_pivot_baseline.style.background_gradient(cmap='viridis', axis=None))
                download_dataframe_csv(req_pivot_baseline, f"whatif_baseline_req_{selected_scenario_name}_{selected_week_str}")

                st.subheader("Simulated Required Staffing")
                st.info("The new staffing requirements after applying your simulation controls.")
                st.dataframe(req_pivot_sim.style.background_gradient(cmap='viridis', axis=None))
                download_dataframe_csv(req_pivot_sim, f"whatif_simulated_req_{selected_scenario_name}_{selected_week_str}")

                st.subheader("Difference (Simulated - Baseline)")
                st.info("The change in required staff for each interval. Positive values (blue) mean more staff are needed; negative values (red) mean fewer staff are needed.")
                st.dataframe(req_diff.style.background_gradient(cmap='RdBu', axis=None))
                download_dataframe_csv(req_diff, f"whatif_req_diff_{selected_scenario_name}_{selected_week_str}")
# --- FOOTER ---
# It's good practice to add a placeholder to avoid content overlapping the footer.
# This one needs to be a bit taller to accommodate the design.
st.markdown("<div style='height: 15rem;'></div>", unsafe_allow_html=True)

# Get the current year for the copyright
current_year = datetime.datetime.now().year

# --- DEFINE THE HTML AND CSS ---
# NOTE: All CSS curly braces are doubled {{ }} to escape them in the f-string.
footer_html = f"""
<style>
@keyframes shimmer {{
    0% {{
        background-position: -1000px 0;
    }}
    100% {{
        background-position: 1000px 0;
    }}
}}

@keyframes slideUp {{
    from {{
        transform: translateY(30px);
        opacity: 0;
    }}
    to {{
        transform: translateY(0);
        opacity: 1;
    }}
}}

:root {{
    --bg-color: #020c1b; /* The darkest blue, almost black */
    --primary-text: #ccd6f6;
    --secondary-text: #8892b0;
    --highlight-color: #64ffda; /* The vibrant mint green */
    --spotlight-color: rgba(100, 255, 218, 0.1); /* Spotlight color with more alpha */
    --footer-font-family: 'SF Mono', 'Fira Code', 'Fira Mono', 'Roboto Mono', monospace;
}}

.footer {{
    font-family: var(--footer-font-family);
    position: fixed;
    bottom: 0;
    left: 0;
    width: 100%;
    background: var(--bg-color);
    color: var(--secondary-text);
    padding: 2rem 3rem;
    z-index: 1000;
    box-shadow: 0 -10px 30px -15px rgba(2, 12, 27, 0.7);
    transition: background 0.3s ease-in-out;

    /* This creates the spotlight effect based on cursor position */
    --mouse-x: 50%;
    --mouse-y: 50%;
    background-image: radial-gradient(
        circle 300px at var(--mouse-x) var(--mouse-y),
        var(--spotlight-color),
        transparent 80%
    );
}}

.footer::before {{
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    width: 100%;
    height: 2px;
    background: linear-gradient(90deg, transparent, var(--highlight-color), transparent);
    animation: shimmer 4s linear infinite;
}}

.footer-content {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    max-width: 1200px;
    margin: 0 auto;
    gap: 1.5rem;
}}

.footer-brand {{
    display: flex;
    align-items: center;
    gap: 1rem;
    animation: slideUp 0.8s cubic-bezier(0.645, 0.045, 0.355, 1) 0.2s forwards;
    opacity: 0;
}}

.footer-logo svg {{
    width: 40px;
    height: 40px;
    fill: var(--highlight-color);
    transition: transform 0.3s ease;
}}
.footer-logo:hover svg {{
    transform: rotate(-10deg) scale(1.1);
}}

.footer-credit {{
    font-size: 0.9rem;
}}
.footer-credit a {{
    color: var(--primary-text);
    text-decoration: none;
    font-weight: bold;
    transition: color 0.3s ease;
}}
.footer-credit a:hover {{
    color: var(--highlight-color);
}}
.footer-copyright {{
    font-size: 0.8rem;
}}

.footer-links {{
    display: flex;
    align-items: center;
    gap: 1.5rem;
}}

.footer-item {{
    opacity: 0;
    animation: slideUp 0.8s cubic-bezier(0.645, 0.045, 0.355, 1) forwards;
}}
.footer-item:nth-child(1) {{ animation-delay: 0.4s; }}
.footer-item:nth-child(2) {{ animation-delay: 0.6s; }}
.footer-item:nth-child(3) {{ animation-delay: 0.8s; }}

.footer-item a {{
    color: var(--secondary-text);
    text-decoration: none;
    display: inline-block;
    position: relative;
    transition: color 0.3s ease, transform 0.3s ease;
}}

/* Animated Underline */
.footer-item a::after {{
    content: '';
    position: absolute;
    width: 100%;
    height: 1px;
    bottom: -4px;
    left: 0;
    background-color: var(--highlight-color);
    transform: scaleX(0);
    transform-origin: bottom right;
    transition: transform 0.3s ease-out;
}}

.footer-item:hover a {{
    color: var(--highlight-color);
    transform: translateY(-3px);
}}
.footer-item:hover a::after {{
    transform: scaleX(1);
    transform-origin: bottom left;
}}

.footer-item svg {{
    width: 24px;
    height: 24px;
    fill: var(--secondary-text);
    margin-right: 0.5rem;
    transition: fill 0.3s ease, transform 0.3s ease;
}}

.footer-item:hover svg {{
    fill: var(--highlight-color);
    transform: translateY(-3px) scale(1.15) rotate(-10deg);
}}

@media (max-width: 768px) {{
    .footer-content {{
        flex-direction: column;
        justify-content: center;
        gap: 2rem;
    }}
}}
</style>

<div class="footer" id="page-footer">
    <div class="footer-content">
        <div class="footer-brand">
            <a href="https://www.linkedin.com/in/ashwin-nair-se/" target="_blank" rel="noopener noreferrer" class="footer-logo" title="Ashwin Nair">
                <!-- 'AN' Initials SVG Logo -->
                <svg viewBox="0 0 512 512" xmlns="http://www.w3.org/2000/svg">
                    <text x="50%" y="60%" dominant-baseline="middle" text-anchor="middle" font-family="var(--footer-font-family)" font-size="300" font-weight="bold" fill="currentColor">
                        AN
                    </text>
                </svg>
            </a>
            <div class="footer-credit">
                <a href="https://www.linkedin.com/in/ashwin-nair-se/" target="_blank" rel="noopener noreferrer">Ashwin Nair</a>
                <div class="footer-copyright">© {current_year} All rights reserved.</div>
            </div>
        </div>
        <div class="footer-links">
            <div class="footer-item">
                <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M0 3v18h24v-18h-24zm21.518 2l-9.518 7.713-9.518-7.713h19.036zm-19.518 14v-11.817l10 8.104 10-8.104v11.817h-20z"/></svg>
                <a href="mailto:19.nair@gmail.com" target="_blank" rel="noopener noreferrer">19.nair@gmail.com</a>
            </div>
            <div class="footer-item">
                <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M20 22.621l-3.521-6.795c-.32-.607-1.045-1.001-1.789-1.001-1.298 0-2.355 1.057-2.355 2.355 0 .285.053.557.147.809l-1.315-.815c-1.631-.998-3.41-2.223-5.26-3.929-1.92-1.764-3.32-3.83-4.225-5.918l-.54-1.21c.214-.082.434-.131.666-.131 1.298 0 2.355-1.057 2.355-2.355 0-.745-.395-1.47-1.002-1.789l-6.793-3.521-2.123 4.086c.038.165.08.336.141.516.484 1.459 1.258 2.846 2.26 4.341 1.096 1.639 2.454 3.193 4.14 4.793s3.472 2.915 5.166 4.013c1.492.973 2.873 1.742 4.33 2.223.181.061.352.103.518.14l4.085-2.122z"/></svg>
                <a href="tel:+918652508240" target="_blank" rel="noopener noreferrer">+91 8652508240</a>
            </div>
        </div>
    </div>
</div>

<script>
    const footer = document.getElementById('page-footer');
    if (footer) {{
        document.addEventListener('mousemove', (e) => {{
            const rect = footer.getBoundingClientRect();
            // Get mouse position relative to the footer element
            const x = e.clientX - rect.left;
            const y = e.clientY - rect.top;
            
            // Set the CSS variables
            footer.style.setProperty('--mouse-x', `${{x}}px`);
            footer.style.setProperty('--mouse-y', `${{y}}px`);
        }});
    }}
</script>
"""

# --- Render the Footer ---
st.markdown(footer_html, unsafe_allow_html=True)
