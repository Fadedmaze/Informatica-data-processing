"""
Narrative summary text template — edit this file to change report wording.

Search tips:
  - Section headings:  SECTION_
  - Line formats:      LINE_
  - Labels/headings:   LABEL_
  - Thresholds:        THRESHOLD_

Placeholders use Python str.format() syntax, e.g. {name}, {current:,.2f}.
"""

# ---------------------------------------------------------------------------
# Section headings (appear once at the start of each report block)
# ---------------------------------------------------------------------------

SECTION_LOG_TYPE = "Log type comparison:"
SECTION_ORGANIZATION = "Organization comparison:"
SECTION_ENVIRONMENT = "Environment comparison:"
SECTION_PROJECT_ANOMALIES = "Project anomalies:"
SECTION_TASK_ANOMALIES = "Task anomalies:"
SECTION_TASK_SPIKES_WEEK = "Task Spikes (Past Week):"
SECTION_TASK_SPIKES_MONTH = "Task Spikes (Month):"
SECTION_TASK_SPIKES_30_DAYS = "Task Spikes (Rolling Month):"

# ---------------------------------------------------------------------------
# Totals (one line per enabled time period)
# ---------------------------------------------------------------------------

LINE_PAST_WEEK_TOTAL = (
    "Past week ({range}): total IPUs {current:,.2f} "
    "-> previous week {previous:,.2f} ({change})"
)
LINE_PAST_MONTH_TOTAL = (
    "Past calendar month ({range}): total IPUs {current:,.2f} "
    "-> previous period {previous:,.2f} ({change})"
)
LINE_PAST_30_DAYS_TOTAL = (
    "Past 30 days ({range}): total IPUs {current:,.2f} "
    "-> previous 30-day period {previous:,.2f} ({change})"
)

# ---------------------------------------------------------------------------
# Dimension blocks (organization, log type, environment)
# ---------------------------------------------------------------------------

# Header for each dimension value (org name, log type, etc.).
# Placeholders: {label} = e.g. "Organization", {name} = e.g. "BYU-Prod"
# Examples:
#   "{label}: {name}"  ->  Organization: BYU-Prod
#   "{name}"           ->  BYU-Prod
DIMENSION_HEADER = "{name}"

# One indented line per time period under each dimension value.
# Placeholders: {period}, {current}, {previous}, {change}
DIMENSION_PERIOD = (
    "  {period}: {current:,.2f} IPUs -> {previous:,.2f} previous ({change})"
)

# When no data exists for a dimension in the selected periods.
DIMENSION_NO_DATA = (
    "No {label_lower}-level data available{filter_note} in the selected period(s)."
)

# Labels passed into dimension builders (used in DIMENSION_HEADER / no-data text).
LABEL_LOG_TYPE = "Log type"
LABEL_ORGANIZATION = "Organization"
LABEL_ENVIRONMENT = "Environment"

# ---------------------------------------------------------------------------
# Anomalies
# ---------------------------------------------------------------------------

LINE_NONE_DETECTED = "- None detected"
LINE_NONE_DETECTED_INDENT = "  - None detected"

LINE_PROJECT_ANOMALY = (
    "- {period}: {org} / {project} "
    "IPUs {previous:.2f} -> {current:.2f} ({change})"
)

LABEL_TASK_CHANGES = "- Changes"
LABEL_TASK_STARTS_STOPS = "- Starts and stops"
LABEL_GOING_UP = "    - Going up"
LABEL_GOING_DOWN = "    - Going down"
LABEL_STARTS = "Starts"
LABEL_STOPS = "Stops"

LINE_TASK_ANOMALY = (
    "      - {org} / {project} / {task} "
    "IPUs {previous:.2f} -> {current:.2f} ({change}), "
    "Runs {runs_previous:d} -> {runs_current:d} ({runs_change})"
)

# Period sub-headings under task anomalies (indented).
LINE_TASK_ANOMALY_PERIOD = "  - {period}"
LINE_TASK_ANOMALY_CHANGE_TYPE = "    - {change_type}"

# ---------------------------------------------------------------------------
# Task spikes
# ---------------------------------------------------------------------------

LINE_SPIKE_NONE = "- No task showed a major spike versus prior baseline."
LINE_SPIKE_SUMMARY = "- {count} task spike(s) flagged. Top items:"
LINE_SPIKE_ITEM = (
    "  - {date}: {task} in {org} / {project} "
    "ran materially above baseline (about {multiplier:.1f}x)."
)

# ---------------------------------------------------------------------------
# Thresholds (control what gets flagged)
# ---------------------------------------------------------------------------

THRESHOLD_PROJECT_IPU_DELTA = 5
THRESHOLD_TASK_IPU_DELTA = 1
THRESHOLD_TASK_RUN_DELTA = 50

# ---------------------------------------------------------------------------
# Format helpers (used by narrative_summary.py — no need to edit below)
# ---------------------------------------------------------------------------


def format_change(value, precision=2):
    """Format a delta with up/down arrow."""
    direction = "↑" if value >= 0 else "↓"
    return f"{direction} {abs(value):,.{precision}f}"


def format_dimension_header(label, name):
    return DIMENSION_HEADER.format(label=label, name=name)


def format_dimension_period(period, current, previous, change):
    return DIMENSION_PERIOD.format(
        period=period,
        current=current,
        previous=previous,
        change=change,
    )


def format_dimension_no_data(label_lower, filter_note=""):
    return DIMENSION_NO_DATA.format(label_lower=label_lower, filter_note=filter_note)
