"""Narrative summary UI and charts for historical analysis."""

from datetime import timedelta

import numpy as np
import pandas as pd
import streamlit as st

from narrative_report import build_narrative_report
from narrative_template import format_change
from processing import LOG_TYPE_MASS_INGESTION


def effective_metrics(df, ipu_factor, mi_ipu_factor, cost_per_ipu):
    """Compute effective IPUs/cost from raw task rows (respects log type factors)."""
    if df is None or df.empty:
        return pd.DataFrame(columns=['org', 'project_name', 'task_name', 'effective_ipus', 'effective_cost'])

    out = df.copy()

    if 'ipus' in out.columns:
        ipus = pd.to_numeric(out['ipus'], errors='coerce')
    else:
        ipus = pd.Series([pd.NA] * len(out), index=out.index)

    if 'metered_value' in out.columns:
        metered = pd.to_numeric(out['metered_value'], errors='coerce').fillna(0)
    else:
        metered = pd.Series([0.0] * len(out), index=out.index)

    if 'cost' in out.columns:
        cost = pd.to_numeric(out['cost'], errors='coerce')
    else:
        cost = pd.Series([pd.NA] * len(out), index=out.index)

    if 'log_type' in out.columns:
        factors = np.where(
            out['log_type'].fillna('').astype(str).str.strip() == LOG_TYPE_MASS_INGESTION,
            mi_ipu_factor,
            ipu_factor,
        )
    else:
        factors = ipu_factor

    out['effective_ipus'] = ipus.fillna(metered * factors).fillna(0)
    out['effective_cost'] = cost.fillna(out['effective_ipus'] * cost_per_ipu).fillna(0)

    for col in ['org', 'project_name', 'task_name']:
        if col not in out.columns:
            out[col] = '(Unknown)'
        out[col] = out[col].fillna('(Unknown)').astype(str)

    return out


def display_narrative_summary(
    *,
    analysis_end,
    log_type,
    log_type_choice,
    fmt_date,
    ipu_factor,
    mi_ipu_factor,
    cost_per_ipu,
    get_daily_stats_by_date_range,
    get_log_type_stats_by_date_range,
    get_org_stats_by_date_range,
    get_environment_stats_by_date_range,
    get_tasks_by_date_range,
    get_task_spikes_for_period,
):
    """Render the Narrative Summary section in historical analysis."""
    _fmt_date = fmt_date

    def _effective_metrics(df):
        return effective_metrics(df, ipu_factor, mi_ipu_factor, cost_per_ipu)

    st.subheader("Narrative Summary")
    st.write(
        "Copyable executive summary with totals, per-org lines, and anomaly bullets. "
        "Edit report wording in `narrative_template.py`."
    )

    compare_full_months = st.checkbox(
        "Compare full previous months",
        value=False,
        key="narrative_compare_full_months",
    )

    st.markdown("**Report options**")

    period_col1, period_col2, period_col3 = st.columns(3)
    with period_col1:
        show_week = st.checkbox(
            "Show past week",
            value=True,
            key="narrative_show_week",
        )
    with period_col2:
        show_month = st.checkbox(
            "Show past calendar month",
            value=False,
            key="narrative_show_month",
        )
    with period_col3:
        show_30_days = st.checkbox(
            "Show past 30 days",
            value=False,
            key="narrative_show_30_days",
        )

    option_col1, option_col2 = st.columns(2)
    with option_col1:
        show_environment = st.checkbox(
            "Show environments",
            value=True,
            key="narrative_show_environment",
            help="Include environment-level comparisons in the narrative summary.",
        )
    with option_col2:
        show_starts_stops = st.checkbox(
            "Show starts and stops",
            value=False,
            key="narrative_show_starts_stops",
            help="Include tasks that started or stopped in the task anomalies section.",
        )

    options_key = "|".join([
        str(show_week),
        str(show_month),
        str(show_30_days),
        str(show_environment),
        str(show_starts_stops),
        str(compare_full_months),
        str(log_type),
        str(analysis_end),
    ])

    selected_periods = [
        period_name for period_name, enabled in [
            ("past week", show_week),
            ("past calendar month", show_month),
            ("past 30 days", show_30_days),
        ] if enabled
    ]

    # (Spike past-weeks control moved below to sit under the Spike Task Charts header)

    generate_narrative = st.button(
        "Generate narrative summary",
        type="primary",
        key="generate_narrative_summary",
        help="Heavy over large history — only runs when you click.",
    )

    if not selected_periods:
        st.info("Turn on at least one time period to generate the narrative summary.")
    elif not generate_narrative:
        cached = st.session_state.get("narrative_report_text")
        cached_options = st.session_state.get("narrative_report_options_key")
        cache_matches_options = cached_options == options_key
        if cached and cache_matches_options:
            st.text_area(
                "Narrative Summary (cached)",
                value=cached,
                height=420,
                key="narrative_summary_cached_view",
            )
            st.caption("Click Generate again to rebuild with the current options.")
        elif cached:
            st.warning(
                "Report options changed since the last generate "
                "(e.g. starts/stops toggle). Click **Generate narrative summary** to rebuild."
            )
        else:
            st.info("Click “Generate narrative summary” to build the report for the selected periods.")

    if selected_periods and generate_narrative:
        report_lines, org_list, task_flags, anchor = build_narrative_report(
            anchor=analysis_end,
            show_week=show_week,
            show_month=show_month,
            show_30_days=show_30_days,
            show_environment=show_environment,
            show_starts_stops=show_starts_stops,
            compare_full_months=compare_full_months,
            log_type=log_type,
            log_type_choice=log_type_choice,
            fmt_date=fmt_date,
            effective_metrics_fn=_effective_metrics,
            get_daily_stats_by_date_range=get_daily_stats_by_date_range,
            get_log_type_stats_by_date_range=get_log_type_stats_by_date_range,
            get_org_stats_by_date_range=get_org_stats_by_date_range,
            get_environment_stats_by_date_range=get_environment_stats_by_date_range,
            get_tasks_by_date_range=get_tasks_by_date_range,
            get_task_spikes_for_period=get_task_spikes_for_period,
        )

        def _normalize_dimension_value(value):
            if pd.isna(value):
                return "(Unknown)"
            text = str(value).strip()
            return text if text else "(Unknown)"

        def _arrow_change(value, precision=2):
            return format_change(value, precision)

        st.text_area("Copy/Paste Report", value="\n".join(report_lines), height=520)
        st.session_state["narrative_report_text"] = "\n".join(report_lines)
        st.session_state["narrative_report_options_key"] = options_key

        # --- Per-organization 6-week weekly trends for screenshotting ---
        try:
            st.divider()
            st.subheader("Org Weekly Trends (last 6 weeks)")

            weeks = 6
            anchor_ts = pd.to_datetime(anchor)
            # align anchor to week start (Monday)
            anchor_week_start = anchor_ts - pd.Timedelta(days=anchor_ts.weekday())
            start_week_start = anchor_week_start - pd.Timedelta(days=7 * (weeks - 1))

            trend_start_iso = start_week_start.date().isoformat()
            trend_end_iso = anchor_ts.date().isoformat()

            trend_tasks = get_tasks_by_date_range(trend_start_iso, trend_end_iso)

            if trend_tasks.empty:
                st.info("No data available for org weekly trends")
            else:
                trend_eff = _effective_metrics(trend_tasks)
                # derive week_start column (Timestamp at week start)
                if 'end_time' in trend_eff.columns:
                    trend_eff['week_start'] = pd.to_datetime(trend_eff['end_time'], errors='coerce').dt.to_period('W').apply(lambda p: p.start_time)
                elif 'start_time' in trend_eff.columns:
                    trend_eff['week_start'] = pd.to_datetime(trend_eff['start_time'], errors='coerce').dt.to_period('W').apply(lambda p: p.start_time)
                else:
                    trend_eff['week_start'] = pd.to_datetime(trend_eff['Run Date'], errors='coerce').dt.to_period('W').apply(lambda p: p.start_time)

                # Build canonical week index
                week_index = pd.date_range(start=start_week_start.normalize(), periods=weeks, freq='7D')

                # Include every org in the trend window (Mass Ingestion is its own org)
                # plus any orgs already listed in the narrative summary.
                trend_orgs = []
                if 'org' in trend_eff.columns and not trend_eff.empty:
                    trend_orgs = [
                        _normalize_dimension_value(v)
                        for v in trend_eff['org'].tolist()
                    ]
                chart_orgs = sorted({
                    *(org_list or []),
                    *trend_orgs,
                })

                # Render org charts in a grid: 3 per row
                for i in range(0, len(chart_orgs), 3):
                    row_orgs = chart_orgs[i:i+3]
                    cols = st.columns(3)
                    for col_idx, org_name in enumerate(row_orgs):
                        with cols[col_idx]:
                            org_df = trend_eff[
                                trend_eff['org'].map(_normalize_dimension_value) == org_name
                            ].copy() if 'org' in trend_eff.columns else pd.DataFrame()
                            if org_df.empty:
                                weekly = pd.DataFrame(
                                    {'total_ipus': [0.0] * len(week_index)},
                                    index=week_index
                                )
                                st.markdown(f"**{org_name}**")
                                st.info("No output in the selected recent-week window.")
                                st.line_chart(weekly['total_ipus'], width='stretch', height=220)
                                continue
                            weekly = org_df.groupby('week_start', dropna=False).agg(total_ipus=('effective_ipus', 'sum')).reset_index()
                            weekly['week_start'] = pd.to_datetime(weekly['week_start'])
                            weekly = weekly.set_index('week_start').reindex(week_index, fill_value=0.0)
                            weekly.index = pd.to_datetime(weekly.index)

                            st.markdown(f"**{org_name}**")
                            st.line_chart(weekly['total_ipus'], width='stretch', height=220)

        except Exception:
            # don't crash the report UI if charting fails
            st.warning('Unable to render org weekly trends')

        # --- New: Plot recent task-level spikes/anomalies as small charts ---
        try:
            st.divider()
            st.subheader('Spike Task Charts (recent weeks)')

            # Show an input here so the user can choose how many past weeks to include
            spike_past_weeks = st.number_input(
                "Past weeks to include for spike charts",
                min_value=1,
                max_value=52,
                value=int(st.session_state.get('spike_past_weeks', 3)),
                step=1,
                help="Enter how many past weeks to include when building spike task charts",
                key="spike_past_weeks_inline",
            )

            # Convert weeks to days for lookback window
            spike_window_days = int(spike_past_weeks) * 7
            spikes = get_task_spikes_for_period(
                end_date=anchor.isoformat(),
                lookback_days=spike_window_days,
                baseline_days=spike_window_days,
                threshold_std=3.0,
                min_baseline_days=5,
                top_n=10,
            )

            if spikes.empty:
                st.info('No recent task spikes to show')
            else:
                # Prefer to show the same tasks and ordering as the narrative "Changes" report
                try:
                    raw_range_start = (anchor - timedelta(days=spike_window_days - 1)).isoformat()
                    raw_range_end = anchor.isoformat()
                    raw_tasks = get_tasks_by_date_range(raw_range_start, raw_range_end)
                except Exception:
                    raw_tasks = pd.DataFrame()

                if raw_tasks.empty:
                    st.info('No raw task rows available for spike charts')
                else:
                    raw_eff = _effective_metrics(raw_tasks)

                    # Build the ordered list from task_flags (if available) else fallback to spikes
                    ordered_tasks = []
                    try:
                        # task_flags exists earlier in this scope and contains delta_ipus/delta_runs
                        if not task_flags.empty:
                            # mirror the report ordering: Changes -> Past Week -> Going up then Going down
                            changes_flags = task_flags[task_flags['change_type'] == 'Changes'].copy()
                            # keep only Past Week entries for narrative context
                            pw = changes_flags[changes_flags['period'] == 'Past Week'] if 'Past Week' in changes_flags['period'].values else changes_flags
                            if not pw.empty:
                                pw = pw.copy()
                                pw['abs_delta'] = pw['delta_ipus'].abs()
                                pw = pw.sort_values('abs_delta', ascending=False)
                                # order: all positive delta_ipus then negative
                                up = pw[pw['delta_ipus'] > 0]
                                down = pw[pw['delta_ipus'] < 0]
                                ordered_df = pd.concat([up, down], ignore_index=True)
                                ordered_tasks = ordered_df[['org', 'project_name', 'task_name', 'delta_ipus', 'delta_runs']].drop_duplicates().to_dict('records')
                    except Exception:
                        ordered_tasks = []

                    # fallback: use spikes list (unique task/project/org order by z_score)
                    if not ordered_tasks:
                        dedup = spikes.drop_duplicates(subset=['task_name', 'org', 'project_name'])
                        dedup = dedup.sort_values(['z_score', 'daily_ipus'], ascending=[False, False])
                        ordered_tasks = dedup[['org', 'project_name', 'task_name']].head(12).to_dict('records')

                    # limit to top_n
                    top_n = 12
                    ordered_tasks = ordered_tasks[:top_n]

                    # Separate entries into IPU-driven and Run-driven lists
                    ipu_entries = []
                    run_entries = []

                    for entry in ordered_tasks:
                        org_n = entry.get('org')
                        project_n = entry.get('project_name')
                        task_name = entry.get('task_name')

                        mask = (
                            raw_eff['task_name'].astype(str) == str(task_name)
                        ) & (
                            raw_eff['org'].astype(str) == str(org_n)
                        ) & (
                            raw_eff['project_name'].astype(str) == str(project_n)
                        )
                        task_hist = raw_eff[mask].copy()
                        if task_hist.empty:
                            continue

                        if 'end_time' in task_hist.columns:
                            task_hist['date'] = pd.to_datetime(task_hist['end_time'], errors='coerce').dt.date
                        elif 'start_time' in task_hist.columns:
                            task_hist['date'] = pd.to_datetime(task_hist['start_time'], errors='coerce').dt.date
                        else:
                            task_hist['date'] = pd.to_datetime(task_hist['Run Date'], errors='coerce').dt.date

                        daily = task_hist.groupby('date').agg(
                            total_ipus=('effective_ipus', 'sum'),
                            runs=('task_name', 'count')
                        ).reset_index()

                        if daily.empty:
                            continue

                        daily['date'] = pd.to_datetime(daily['date'])
                        daily = daily.set_index('date').sort_index()

                        # heuristics to decide run vs ipu
                        delta_runs = entry.get('delta_runs', None)

                        mean_ipu = daily['total_ipus'].mean() if 'total_ipus' in daily.columns else 0.0
                        peak_ipu = daily['total_ipus'].max() if 'total_ipus' in daily.columns else 0.0
                        mult_ipu = (peak_ipu / mean_ipu) if mean_ipu > 0 else (peak_ipu if peak_ipu > 0 else 0.0)

                        mean_runs = daily['runs'].mean() if 'runs' in daily.columns else 0.0
                        peak_runs = daily['runs'].max() if 'runs' in daily.columns else 0.0
                        mult_runs = (peak_runs / mean_runs) if mean_runs > 0 else (peak_runs if peak_runs > 0 else 0.0)

                        choose_runs = False
                        if delta_runs is not None and abs(delta_runs) > 50:
                            choose_runs = True
                        elif mult_runs > mult_ipu:
                            choose_runs = True

                        record = {
                            'org': org_n,
                            'project_name': project_n,
                            'task_name': task_name,
                            'daily': daily,
                            'choose_runs': choose_runs,
                            'delta_ipus': entry.get('delta_ipus', None),
                            'delta_runs': entry.get('delta_runs', None),
                        }

                        if choose_runs:
                            run_entries.append(record)
                        else:
                            ipu_entries.append(record)

                    def _split_and_sort(entries, metric_key, precision=2):
                        up = []
                        down = []
                        for rec in entries:
                            # Prefer explicit delta from task_flags if available
                            spike_val = None
                            if metric_key == 'total_ipus' and rec.get('delta_ipus') is not None:
                                spike_val = float(rec.get('delta_ipus'))
                            if metric_key == 'runs' and rec.get('delta_runs') is not None:
                                spike_val = float(rec.get('delta_runs'))

                            daily = rec['daily']
                            series = daily[metric_key]
                            if spike_val is None:
                                peak_val = series.max()
                                mean_val = series.mean()
                                spike_val = peak_val - mean_val
                            rec['spike_val'] = spike_val
                            # classify by spike sign
                            if spike_val >= 0:
                                up.append(rec)
                            else:
                                down.append(rec)
                        # sort by absolute spike magnitude desc
                        up = sorted(up, key=lambda r: abs(r.get('spike_val', 0)), reverse=True)
                        down = sorted(down, key=lambda r: abs(r.get('spike_val', 0)), reverse=True)
                        return up, down

                    # Render IPU-driven spikes first, split into Going Up / Going Down (collapsible)
                    if ipu_entries:
                        with st.expander('IPU-driven Spike Charts', expanded=False):
                            ipu_up, ipu_down = _split_and_sort(ipu_entries, 'total_ipus', precision=2)

                            if ipu_up:
                                with st.expander('Going Up', expanded=False):
                                    for i in range(0, len(ipu_up), 3):
                                        row = ipu_up[i:i+3]
                                        cols = st.columns(3)
                                        for col_idx, rec in enumerate(row):
                                            with cols[col_idx]:
                                                org_n = rec['org']
                                                project_n = rec['project_name']
                                                task_name = rec['task_name']
                                                daily = rec['daily']
                                                series = daily['total_ipus']
                                                chart_title = f"{task_name} — {org_n} / {project_n} (IPUs)"
                                                st.markdown(f"**{chart_title}**")
                                                spike_val = rec.get('spike_val', 0)
                                                st.metric('Spike', _arrow_change(spike_val, precision=2), delta='(compared to last week)')
                                                st.line_chart(series, width='stretch', height=260)

                            if ipu_down:
                                with st.expander('Going Down', expanded=False):
                                    for i in range(0, len(ipu_down), 3):
                                        row = ipu_down[i:i+3]
                                        cols = st.columns(3)
                                        for col_idx, rec in enumerate(row):
                                            with cols[col_idx]:
                                                org_n = rec['org']
                                                project_n = rec['project_name']
                                                task_name = rec['task_name']
                                                daily = rec['daily']
                                                series = daily['total_ipus']
                                                chart_title = f"{task_name} — {org_n} / {project_n} (IPUs)"
                                                st.markdown(f"**{chart_title}**")
                                                spike_val = rec.get('spike_val', 0)
                                                st.metric('Spike', _arrow_change(spike_val, precision=2), delta='(compared to last week)')
                                                st.line_chart(series, width='stretch', height=260)

                    # Then render Run-driven spikes, split into Going Up / Going Down
                    if run_entries:
                        with st.expander('Run-driven Spike Charts', expanded=False):
                            run_up, run_down = _split_and_sort(run_entries, 'runs', precision=0)

                            if run_up:
                                with st.expander('Going Up', expanded=False):
                                    for i in range(0, len(run_up), 3):
                                        row = run_up[i:i+3]
                                        cols = st.columns(3)
                                        for col_idx, rec in enumerate(row):
                                            with cols[col_idx]:
                                                org_n = rec['org']
                                                project_n = rec['project_name']
                                                task_name = rec['task_name']
                                                daily = rec['daily']
                                                series = daily['runs']
                                                chart_title = f"{task_name} — {org_n} / {project_n} (Runs)"
                                                st.markdown(f"**{chart_title}**")
                                                spike_val = rec.get('spike_val', 0)
                                                st.metric('Spike', _arrow_change(spike_val, precision=0), delta='(compared to last week)')
                                                st.line_chart(series, width='stretch', height=260)

                            if run_down:
                                with st.expander('Going Down', expanded=False):
                                    for i in range(0, len(run_down), 3):
                                        row = run_down[i:i+3]
                                        cols = st.columns(3)
                                        for col_idx, rec in enumerate(row):
                                            with cols[col_idx]:
                                                org_n = rec['org']
                                                project_n = rec['project_name']
                                                task_name = rec['task_name']
                                                daily = rec['daily']
                                                series = daily['runs']
                                                chart_title = f"{task_name} — {org_n} / {project_n} (Runs)"
                                                st.markdown(f"**{chart_title}**")
                                                spike_val = rec.get('spike_val', 0)
                                                st.metric('Spike', _arrow_change(spike_val, precision=0), delta='(compared to last week)')
                                                st.line_chart(series, width='stretch', height=260)

        except Exception:
            st.warning('Unable to render spike charts')
