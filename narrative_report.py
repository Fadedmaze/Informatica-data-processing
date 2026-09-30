"""Assemble narrative summary report text from data using narrative_template."""

from calendar import monthrange
from datetime import timedelta

import pandas as pd

import narrative_template as T


def build_narrative_report(
    *,
    anchor,
    show_week,
    show_month,
    show_30_days,
    show_environment,
    show_starts_stops,
    compare_full_months,
    log_type,
    log_type_choice,
    fmt_date,
    effective_metrics_fn,
    get_daily_stats_by_date_range,
    get_log_type_stats_by_date_range,
    get_org_stats_by_date_range,
    get_environment_stats_by_date_range,
    get_tasks_by_date_range,
    get_task_spikes_for_period,
):
    """Build the copy-paste narrative report lines."""

    def _month_to_date_ranges(anchor_date):
        cur_start = anchor_date.replace(day=1)
        cur_end = anchor_date
        prev_month = (cur_start.month - 1) or 12
        prev_year = cur_start.year - 1 if cur_start.month == 1 else cur_start.year
        prev_start = cur_start.replace(year=prev_year, month=prev_month, day=1)
        last_day_prev = monthrange(prev_year, prev_month)[1]
        prev_end_day = min(anchor_date.day, last_day_prev)
        prev_end = prev_start.replace(day=prev_end_day)
        return cur_start, cur_end, prev_start, prev_end

    if compare_full_months:
        if anchor.month == 1:
            cur_month = 12
            cur_year = anchor.year - 1
        else:
            cur_month = anchor.month - 1
            cur_year = anchor.year
        cur_start = anchor.replace(year=cur_year, month=cur_month, day=1)
        cur_last = monthrange(cur_year, cur_month)[1]
        cur_end = cur_start.replace(day=cur_last)
        if cur_month == 1:
            prev_month = 12
            prev_year = cur_year - 1
        else:
            prev_month = cur_month - 1
            prev_year = cur_year
        prev_start = cur_start.replace(year=prev_year, month=prev_month, day=1)
        prev_last = monthrange(prev_year, prev_month)[1]
        prev_end = prev_start.replace(day=prev_last)
    else:
        cur_start, cur_end, prev_start, prev_end = _month_to_date_ranges(anchor)

    week_cur_end = anchor
    week_cur_start = anchor - timedelta(days=6)
    week_prev_end = week_cur_start - timedelta(days=1)
    week_prev_start = week_prev_end - timedelta(days=6)

    d30_cur_end = anchor
    d30_cur_start = anchor - timedelta(days=29)
    d30_prev_end = d30_cur_start - timedelta(days=1)
    d30_prev_start = d30_prev_end - timedelta(days=29)

    def _total_ipus_for_range(start, end):
        df = get_daily_stats_by_date_range(start.isoformat(), end.isoformat())
        return float(df["total_ipus"].sum()) if not df.empty else 0.0

    total_week = _total_ipus_for_range(week_cur_start, week_cur_end)
    total_week_prev = _total_ipus_for_range(week_prev_start, week_prev_end)
    total_30 = _total_ipus_for_range(d30_cur_start, d30_cur_end)
    total_30_prev = _total_ipus_for_range(d30_prev_start, d30_prev_end)
    total_mtd = _total_ipus_for_range(cur_start, cur_end)
    total_mtd_prev = _total_ipus_for_range(prev_start, prev_end)

    week_range_str = (
        f"{week_cur_start.strftime('%b')} {week_cur_start.day} - "
        f"{week_cur_end.strftime('%b')} {week_cur_end.day}"
    )
    mtd_range_str = (
        f"{cur_start.strftime('%b')} {cur_start.day} - "
        f"{cur_end.strftime('%b')} {cur_end.day}"
    )
    d30_range_str = (
        f"{d30_cur_start.strftime('%b')} {d30_cur_start.day} - "
        f"{d30_cur_end.strftime('%b')} {d30_cur_end.day}"
    )

    def _normalize_dimension_value(value):
        if pd.isna(value):
            return "(Unknown)"
        text = str(value).strip()
        return text if text else "(Unknown)"

    def _build_dimension_summary_lines(label, dimension_key, stats_fn):
        period_defs = []
        if show_week:
            period_defs.append(
                ("Past week", week_cur_start, week_cur_end, week_prev_start, week_prev_end)
            )
        if show_month:
            period_defs.append(
                ("Past calendar month", cur_start, cur_end, prev_start, prev_end)
            )
        if show_30_days:
            period_defs.append(
                ("Past 30 days", d30_cur_start, d30_cur_end, d30_prev_start, d30_prev_end)
            )

        period_frames = {}
        for _, cur_s, cur_e, prev_s, prev_e in period_defs:
            period_frames[("cur", cur_s, cur_e)] = stats_fn(cur_s.isoformat(), cur_e.isoformat())
            period_frames[("prev", prev_s, prev_e)] = stats_fn(prev_s.isoformat(), prev_e.isoformat())

        dimension_values = sorted({
            _normalize_dimension_value(value)
            for frame in period_frames.values()
            if frame is not None and not frame.empty and dimension_key in frame.columns
            for value in frame[dimension_key].tolist()
        })

        lines = []
        if not dimension_values:
            filter_note = f" for log type '{log_type_choice}'" if log_type else ""
            lines.append(T.format_dimension_no_data(label.lower(), filter_note))
            lines.append("")
            return lines, []

        for dimension_name in dimension_values:
            lines.append(T.format_dimension_header(label, dimension_name))
            for period_name, cur_s, cur_e, prev_s, prev_e in period_defs:
                cur_df = period_frames[("cur", cur_s, cur_e)]
                prev_df = period_frames[("prev", prev_s, prev_e)]

                if cur_df.empty or dimension_key not in cur_df.columns:
                    cur_val = 0.0
                else:
                    cur_mask = (
                        cur_df[dimension_key].map(_normalize_dimension_value) == dimension_name
                    )
                    cur_val = (
                        float(cur_df.loc[cur_mask, "total_ipus"].sum())
                        if cur_mask.any()
                        else 0.0
                    )

                if prev_df.empty or dimension_key not in prev_df.columns:
                    prev_val = 0.0
                else:
                    prev_mask = (
                        prev_df[dimension_key].map(_normalize_dimension_value) == dimension_name
                    )
                    prev_val = (
                        float(prev_df.loc[prev_mask, "total_ipus"].sum())
                        if prev_mask.any()
                        else 0.0
                    )

                lines.append(
                    T.format_dimension_period(
                        period_name,
                        cur_val,
                        prev_val,
                        T.format_change(cur_val - prev_val),
                    )
                )

            lines.append("")

        return lines, dimension_values

    report_lines = []

    if show_week:
        report_lines.append(
            T.LINE_PAST_WEEK_TOTAL.format(
                range=week_range_str,
                current=total_week,
                previous=total_week_prev,
                change=T.format_change(total_week - total_week_prev),
            )
        )

    if show_month:
        report_lines.append(
            T.LINE_PAST_MONTH_TOTAL.format(
                range=mtd_range_str,
                current=total_mtd,
                previous=total_mtd_prev,
                change=T.format_change(total_mtd - total_mtd_prev),
            )
        )

    if show_30_days:
        report_lines.append(
            T.LINE_PAST_30_DAYS_TOTAL.format(
                range=d30_range_str,
                current=total_30,
                previous=total_30_prev,
                change=T.format_change(total_30 - total_30_prev),
            )
        )

    report_lines.append("")

    report_lines.append(T.SECTION_LOG_TYPE)
    log_lines, _log_list = _build_dimension_summary_lines(
        T.LABEL_LOG_TYPE,
        "log_type",
        lambda s, e: get_log_type_stats_by_date_range(s, e),
    )
    report_lines.extend(log_lines)

    report_lines.append(T.SECTION_ORGANIZATION)
    org_lines, org_list = _build_dimension_summary_lines(
        T.LABEL_ORGANIZATION,
        "org",
        get_org_stats_by_date_range,
    )
    report_lines.extend(org_lines)

    if show_environment:
        report_lines.append(T.SECTION_ENVIRONMENT)
        env_lines, _env_list = _build_dimension_summary_lines(
            T.LABEL_ENVIRONMENT,
            "environment",
            get_environment_stats_by_date_range,
        )
        report_lines.extend(env_lines)

    def _detect_project_anomalies(cur_s, cur_e, prev_s, prev_e, label):
        cur_eff = effective_metrics_fn(get_tasks_by_date_range(cur_s.isoformat(), cur_e.isoformat()))
        prev_eff = effective_metrics_fn(get_tasks_by_date_range(prev_s.isoformat(), prev_e.isoformat()))
        cur_grp = (
            cur_eff.groupby(["org", "project_name"])
            .agg(total_ipus=("effective_ipus", "sum"))
            .reset_index()
        )
        prev_grp = (
            prev_eff.groupby(["org", "project_name"])
            .agg(total_ipus=("effective_ipus", "sum"))
            .reset_index()
        )
        merged = cur_grp.merge(
            prev_grp, on=["org", "project_name"], how="outer", suffixes=("_cur", "_prev")
        ).fillna(0)
        merged["delta_ipus"] = merged["total_ipus_cur"] - merged["total_ipus_prev"]
        flagged = merged[merged["delta_ipus"].abs() >= T.THRESHOLD_PROJECT_IPU_DELTA].copy()
        flagged["period"] = label
        return flagged

    proj_frames = []
    if show_week:
        proj_frames.append(
            _detect_project_anomalies(
                week_cur_start, week_cur_end, week_prev_start, week_prev_end, "Past Week"
            )
        )
    if show_month:
        proj_frames.append(
            _detect_project_anomalies(cur_start, cur_end, prev_start, prev_end, "Month-to-date")
        )
    if show_30_days:
        proj_frames.append(
            _detect_project_anomalies(
                d30_cur_start, d30_cur_end, d30_prev_start, d30_prev_end, "Past 30 Days"
            )
        )

    proj_flags = pd.concat(proj_frames, ignore_index=True, sort=False) if proj_frames else pd.DataFrame()

    report_lines.append(T.SECTION_PROJECT_ANOMALIES)
    if proj_flags.empty:
        report_lines.append(T.LINE_NONE_DETECTED)
    else:
        period_order = {"Past Week": 0, "Month-to-date": 1, "Past 30 Days": 2}
        proj_flags = proj_flags.copy()
        proj_flags["period_rank"] = proj_flags["period"].map(period_order).fillna(99).astype(int)
        proj_flags["abs_delta"] = proj_flags["delta_ipus"].abs()
        for _, row in proj_flags.sort_values(
            ["period_rank", "abs_delta"], ascending=[True, False]
        ).iterrows():
            report_lines.append(
                T.LINE_PROJECT_ANOMALY.format(
                    period=row["period"],
                    org=row["org"],
                    project=row["project_name"],
                    previous=row["total_ipus_prev"],
                    current=row["total_ipus_cur"],
                    change=T.format_change(row["delta_ipus"]),
                )
            )

    report_lines.append("")

    def _detect_task_anomalies(cur_s, cur_e, prev_s, prev_e, label):
        cur_eff = effective_metrics_fn(get_tasks_by_date_range(cur_s.isoformat(), cur_e.isoformat()))
        prev_eff = effective_metrics_fn(get_tasks_by_date_range(prev_s.isoformat(), prev_e.isoformat()))
        cur_grp = (
            cur_eff.groupby(["org", "project_name", "task_name"])
            .agg(task_count=("task_name", "count"), total_ipus=("effective_ipus", "sum"))
            .reset_index()
        )
        prev_grp = (
            prev_eff.groupby(["org", "project_name", "task_name"])
            .agg(task_count=("task_name", "count"), total_ipus=("effective_ipus", "sum"))
            .reset_index()
        )
        merged = cur_grp.merge(
            prev_grp,
            on=["org", "project_name", "task_name"],
            how="outer",
            suffixes=("_cur", "_prev"),
        ).fillna(0)
        merged["delta_ipus"] = merged["total_ipus_cur"] - merged["total_ipus_prev"]
        merged["delta_runs"] = merged["task_count_cur"] - merged["task_count_prev"]
        flagged = merged[
            (merged["delta_ipus"].abs() >= T.THRESHOLD_TASK_IPU_DELTA)
            | (merged["delta_runs"].abs() > T.THRESHOLD_TASK_RUN_DELTA)
        ].copy()
        flagged["period"] = label
        return flagged

    task_frames = []
    if show_week:
        task_frames.append(
            _detect_task_anomalies(
                week_cur_start, week_cur_end, week_prev_start, week_prev_end, "Past Week"
            )
        )
    if show_month:
        task_frames.append(
            _detect_task_anomalies(cur_start, cur_end, prev_start, prev_end, "Month-to-date")
        )
    if show_30_days:
        task_frames.append(
            _detect_task_anomalies(
                d30_cur_start, d30_cur_end, d30_prev_start, d30_prev_end, "Past 30 Days"
            )
        )
    task_flags = pd.concat(task_frames, ignore_index=True, sort=False) if task_frames else pd.DataFrame()

    report_lines.append(T.SECTION_TASK_ANOMALIES)
    if task_flags.empty:
        report_lines.append(T.LINE_NONE_DETECTED)
    else:
        task_flags = task_flags.copy()
        task_flags["change_type"] = task_flags.apply(
            lambda r: (
                T.LABEL_STARTS
                if (r["total_ipus_prev"] <= 0 and r["total_ipus_cur"] > 0)
                else T.LABEL_STOPS
                if (r["total_ipus_prev"] > 0 and r["total_ipus_cur"] <= 0)
                else "Changes"
            ),
            axis=1,
        )
        task_flags["abs_delta"] = task_flags["delta_ipus"].abs()
        report_lines.append(T.LABEL_TASK_CHANGES)
        changes_flags = task_flags[task_flags["change_type"] == "Changes"]
        if changes_flags.empty:
            report_lines.append(T.LINE_NONE_DETECTED_INDENT)
        else:
            for period_name in ["Past Week", "Month-to-date", "Past 30 Days"]:
                if (
                    (period_name == "Past Week" and not show_week)
                    or (period_name == "Month-to-date" and not show_month)
                    or (period_name == "Past 30 Days" and not show_30_days)
                ):
                    continue
                period_flags = changes_flags[
                    changes_flags["period"] == period_name
                ].sort_values("abs_delta", ascending=False)
                if period_flags.empty:
                    continue
                report_lines.append(T.LINE_TASK_ANOMALY_PERIOD.format(period=period_name))
                up_flags = period_flags[period_flags["delta_ipus"] > 0].sort_values(
                    "delta_ipus", ascending=False
                )
                down_flags = period_flags[period_flags["delta_ipus"] < 0].sort_values(
                    "delta_ipus", ascending=True
                )
                if not up_flags.empty:
                    report_lines.append(T.LABEL_GOING_UP)
                    for _, row in up_flags.iterrows():
                        report_lines.append(
                            T.LINE_TASK_ANOMALY.format(
                                org=row["org"],
                                project=row["project_name"],
                                task=row["task_name"],
                                previous=row["total_ipus_prev"],
                                current=row["total_ipus_cur"],
                                change=T.format_change(row["delta_ipus"]),
                                runs_previous=int(row["task_count_prev"]),
                                runs_current=int(row["task_count_cur"]),
                                runs_change=T.format_change(row["delta_runs"], precision=0),
                            )
                        )
                if not down_flags.empty:
                    report_lines.append(T.LABEL_GOING_DOWN)
                    for _, row in down_flags.iterrows():
                        report_lines.append(
                            T.LINE_TASK_ANOMALY.format(
                                org=row["org"],
                                project=row["project_name"],
                                task=row["task_name"],
                                previous=row["total_ipus_prev"],
                                current=row["total_ipus_cur"],
                                change=T.format_change(row["delta_ipus"]),
                                runs_previous=int(row["task_count_prev"]),
                                runs_current=int(row["task_count_cur"]),
                                runs_change=T.format_change(row["delta_runs"], precision=0),
                            )
                        )

        report_lines.append("")
        if show_starts_stops:
            report_lines.append(T.LABEL_TASK_STARTS_STOPS)
            starts_stops_flags = task_flags[task_flags["change_type"].isin([T.LABEL_STARTS, T.LABEL_STOPS])]
            if starts_stops_flags.empty:
                report_lines.append(T.LINE_NONE_DETECTED_INDENT)
            else:
                for period_name in ["Past Week", "Month-to-date", "Past 30 Days"]:
                    if (
                        (period_name == "Past Week" and not show_week)
                        or (period_name == "Month-to-date" and not show_month)
                        or (period_name == "Past 30 Days" and not show_30_days)
                    ):
                        continue
                    period_flags = starts_stops_flags[
                        starts_stops_flags["period"] == period_name
                    ].sort_values("abs_delta", ascending=False)
                    if period_flags.empty:
                        continue
                    report_lines.append(T.LINE_TASK_ANOMALY_PERIOD.format(period=period_name))
                    for change_type in [T.LABEL_STARTS, T.LABEL_STOPS]:
                        change_flags = period_flags[period_flags["change_type"] == change_type]
                        if change_flags.empty:
                            continue
                        report_lines.append(
                            T.LINE_TASK_ANOMALY_CHANGE_TYPE.format(change_type=change_type)
                        )
                        for _, row in change_flags.iterrows():
                            report_lines.append(
                                T.LINE_TASK_ANOMALY.format(
                                    org=row["org"],
                                    project=row["project_name"],
                                    task=row["task_name"],
                                    previous=row["total_ipus_prev"],
                                    current=row["total_ipus_cur"],
                                    change=T.format_change(row["delta_ipus"]),
                                    runs_previous=int(row["task_count_prev"]),
                                    runs_current=int(row["task_count_cur"]),
                                    runs_change=T.format_change(row["delta_runs"], precision=0),
                                )
                            )

    def _add_task_spikes_section(title, end_date, lookback_days, baseline_days):
        report_lines.append("")
        report_lines.append(title)
        spikes = get_task_spikes_for_period(
            end_date=end_date,
            lookback_days=lookback_days,
            baseline_days=baseline_days,
            threshold_std=3.0,
            min_baseline_days=5,
            top_n=10,
        )
        if spikes.empty:
            report_lines.append(T.LINE_SPIKE_NONE)
            return

        report_lines.append(T.LINE_SPIKE_SUMMARY.format(count=len(spikes)))
        for _, row in spikes.head(5).iterrows():
            report_lines.append(
                T.LINE_SPIKE_ITEM.format(
                    date=fmt_date(pd.to_datetime(row["task_date"]).date()),
                    task=row["task_name"],
                    org=row["org"],
                    project=row["project_name"],
                    multiplier=row["multiplier_vs_baseline"],
                )
            )

    if show_week:
        _add_task_spikes_section(T.SECTION_TASK_SPIKES_WEEK, anchor.isoformat(), 7, 7)

    if show_month:
        month_window_days = (cur_end - cur_start).days + 1
        _add_task_spikes_section(
            T.SECTION_TASK_SPIKES_MONTH, cur_end.isoformat(), month_window_days, month_window_days
        )

    if show_30_days:
        _add_task_spikes_section(T.SECTION_TASK_SPIKES_30_DAYS, anchor.isoformat(), 30, 30)

    return report_lines, org_list, task_flags, anchor
