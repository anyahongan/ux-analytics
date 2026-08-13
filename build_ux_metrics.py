from pathlib import Path

import pandas as pd
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


PROJECT_ROOT = Path(__file__).resolve().parent

INPUT_FILE = (
    PROJECT_ROOT
    / "data"
    / "parsed"
    / "master_parsed.xlsx"
)

OUTPUT_FILE = (
    PROJECT_ROOT
    / "data"
    / "parsed"
    / "ux_metrics.xlsx"
)


def require_columns(
    dataframe: pd.DataFrame,
    sheet_name: str,
    required_columns: list[str],
) -> None:
    """
    Stop the script with a clear message when an expected column is missing.
    """
    missing_columns = [
        column
        for column in required_columns
        if column not in dataframe.columns
    ]

    if missing_columns:
        raise ValueError(
            f"The '{sheet_name}' sheet is missing these columns: "
            f"{missing_columns}"
        )


def prepare_datetime_column(
    dataframe: pd.DataFrame,
    column_name: str,
) -> None:
    """
    Convert a column into pandas datetime values when it exists.
    """
    if column_name in dataframe.columns:
        dataframe[column_name] = pd.to_datetime(
            dataframe[column_name],
            errors="coerce",
        )


def make_excel_safe(
    dataframe: pd.DataFrame,
) -> pd.DataFrame:
    """
    Remove timezone metadata from datetime columns before Excel export.
    """
    safe_dataframe = dataframe.copy()

    for column_name in safe_dataframe.columns:
        column = safe_dataframe[column_name]

        if isinstance(column.dtype, pd.DatetimeTZDtype):
            safe_dataframe[column_name] = column.dt.tz_localize(None)

    return safe_dataframe


def auto_fit_columns(writer: pd.ExcelWriter) -> None:
    """
    Resize Excel columns based on their contents.
    """
    for worksheet in writer.book.worksheets:
        for column_cells in worksheet.columns:
            maximum_length = 0
            column_number = column_cells[0].column
            column_letter = get_column_letter(column_number)

            for cell in column_cells:
                if cell.value is None:
                    continue

                maximum_length = max(
                    maximum_length,
                    len(str(cell.value)),
                )

            worksheet.column_dimensions[column_letter].width = min(
                maximum_length + 2,
                55,
            )

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        for cell in worksheet[1]:
            cell.font = Font(bold=True)


def main() -> None:
    print("Reading parsed workbook...")

    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Could not find: {INPUT_FILE.resolve()}"
        )

    sessions_df = pd.read_excel(
        INPUT_FILE,
        sheet_name="Sessions",
    )

    events_df = pd.read_excel(
        INPUT_FILE,
        sheet_name="Events",
    )

    fields_df = pd.read_excel(
        INPUT_FILE,
        sheet_name="Fields",
    )

    quality_df = pd.read_excel(
        INPUT_FILE,
        sheet_name="Data Quality",
    )

    require_columns(
        sessions_df,
        "Sessions",
        [
            "shortIntakeId",
            "formType",
            "startedAt",
            "submittedAt",
            "completionSeconds",
            "completionMinutes",
        ],
    )

    require_columns(
        events_df,
        "Events",
        [
            "shortIntakeId",
            "eventOrder",
            "linkId",
            "eventTime",
        ],
    )

    require_columns(
        fields_df,
        "Fields",
        [
            "shortIntakeId",
            "fieldId",
        ],
    )

    prepare_datetime_column(
        sessions_df,
        "startedAt",
    )

    prepare_datetime_column(
        sessions_df,
        "submittedAt",
    )

    prepare_datetime_column(
        events_df,
        "eventTime",
    )

    if "firstChangedAt" in fields_df.columns:
        prepare_datetime_column(
            fields_df,
            "firstChangedAt",
        )

    # --------------------------------------------------
    # Clean event ordering
    # --------------------------------------------------
    events_df["eventOrder"] = pd.to_numeric(
        events_df["eventOrder"],
        errors="coerce",
    )

    events_df = events_df.sort_values(
        by=[
            "shortIntakeId",
            "eventOrder",
            "eventTime",
        ],
        na_position="last",
    )

    # --------------------------------------------------
    # Event metrics by session
    # --------------------------------------------------
    event_counts = (
        events_df.groupby("shortIntakeId")
        .size()
        .rename("numberOfEvents")
    )

    unique_event_fields = (
        events_df.groupby("shortIntakeId")["linkId"]
        .nunique(dropna=True)
        .rename("uniqueEventFields")
    )

    repeated_event_count = (
        events_df.groupby(
            [
                "shortIntakeId",
                "linkId",
            ]
        )
        .size()
        .sub(1)
        .clip(lower=0)
        .groupby(level=0)
        .sum()
        .rename("repeatEventCount")
    )

    # --------------------------------------------------
    # First event, last event, and event path
    # --------------------------------------------------
    event_details = []

    for short_intake_id, group in events_df.groupby(
        "shortIntakeId",
        sort=False,
    ):
        valid_links = (
            group["linkId"]
            .dropna()
            .astype(str)
            .tolist()
        )

        event_details.append(
            {
                "shortIntakeId": short_intake_id,
                "firstEvent": (
                    valid_links[0]
                    if valid_links
                    else None
                ),
                "lastEvent": (
                    valid_links[-1]
                    if valid_links
                    else None
                ),
                "eventPath": " → ".join(valid_links),
            }
        )

    event_details_df = pd.DataFrame(event_details)

    # --------------------------------------------------
    # Field metrics by session
    # --------------------------------------------------
    field_counts = (
        fields_df.groupby("shortIntakeId")["fieldId"]
        .nunique(dropna=True)
        .rename("numberOfFields")
    )

    # --------------------------------------------------
    # Build Session Summary
    # --------------------------------------------------
    session_summary_df = sessions_df.copy()

    session_summary_df = session_summary_df.merge(
        event_counts,
        on="shortIntakeId",
        how="left",
    )

    session_summary_df = session_summary_df.merge(
        unique_event_fields,
        on="shortIntakeId",
        how="left",
    )

    session_summary_df = session_summary_df.merge(
        repeated_event_count,
        on="shortIntakeId",
        how="left",
    )

    session_summary_df = session_summary_df.merge(
        field_counts,
        on="shortIntakeId",
        how="left",
    )

    if not event_details_df.empty:
        session_summary_df = session_summary_df.merge(
            event_details_df,
            on="shortIntakeId",
            how="left",
        )
    else:
        session_summary_df["firstEvent"] = None
        session_summary_df["lastEvent"] = None
        session_summary_df["eventPath"] = None

    numeric_fill_columns = [
        "numberOfEvents",
        "uniqueEventFields",
        "repeatEventCount",
        "numberOfFields",
    ]

    for column_name in numeric_fill_columns:
        session_summary_df[column_name] = (
            session_summary_df[column_name]
            .fillna(0)
            .astype(int)
        )

    session_summary_df["sessionStatus"] = "Completed"

    session_summary_df.loc[
        session_summary_df["submittedAt"].isna(),
        "sessionStatus",
    ] = "Incomplete"

    session_summary_df["completionSeconds"] = (
        pd.to_numeric(
            session_summary_df["completionSeconds"],
            errors="coerce",
        ).round(2)
    )

    session_summary_df["completionMinutes"] = (
        pd.to_numeric(
            session_summary_df["completionMinutes"],
            errors="coerce",
        ).round(2)
    )

    preferred_session_columns = [
        "shortIntakeId",
        "shortRowKey",
        "formType",
        "sessionStatus",
        "startedAt",
        "submittedAt",
        "completionSeconds",
        "completionMinutes",
        "numberOfEvents",
        "uniqueEventFields",
        "repeatEventCount",
        "numberOfFields",
        "firstEvent",
        "lastEvent",
        "eventPath",
        "submissionType",
        "isEvent",
    ]

    existing_preferred_columns = [
        column
        for column in preferred_session_columns
        if column in session_summary_df.columns
    ]

    remaining_session_columns = [
        column
        for column in session_summary_df.columns
        if column not in existing_preferred_columns
    ]

    session_summary_df = session_summary_df[
        existing_preferred_columns
        + remaining_session_columns
    ]

    # --------------------------------------------------
    # Build Session Paths
    # --------------------------------------------------
    session_paths_df = session_summary_df[
        [
            column
            for column in [
                "shortIntakeId",
                "formType",
                "numberOfEvents",
                "repeatEventCount",
                "firstEvent",
                "lastEvent",
                "eventPath",
            ]
            if column in session_summary_df.columns
        ]
    ].copy()

    # --------------------------------------------------
    # Build Question Summary
    # --------------------------------------------------
    event_question_summary = (
        events_df.groupby("linkId", dropna=False)
        .agg(
            sessionsReached=(
                "shortIntakeId",
                "nunique",
            ),
            totalRecordedEvents=(
                "shortIntakeId",
                "size",
            ),
            averageEventOrder=(
                "eventOrder",
                "mean",
            ),
            medianEventOrder=(
                "eventOrder",
                "median",
            ),
        )
        .reset_index()
        .rename(columns={"linkId": "questionId"})
    )

    repeated_visits_by_question = (
        events_df.groupby(
            [
                "shortIntakeId",
                "linkId",
            ]
        )
        .size()
        .sub(1)
        .clip(lower=0)
        .groupby(level=1)
        .sum()
        .rename("repeatVisits")
        .reset_index()
        .rename(columns={"linkId": "questionId"})
    )

    final_event_counts = (
        event_details_df.groupby("lastEvent")
        .size()
        .rename("sessionsEndingHere")
        .reset_index()
        .rename(columns={"lastEvent": "questionId"})
        if not event_details_df.empty
        else pd.DataFrame(
            columns=[
                "questionId",
                "sessionsEndingHere",
            ]
        )
    )

    question_summary_df = event_question_summary.merge(
        repeated_visits_by_question,
        on="questionId",
        how="left",
    )

    question_summary_df = question_summary_df.merge(
        final_event_counts,
        on="questionId",
        how="left",
    )

    field_session_counts = (
        fields_df.groupby("fieldId")["shortIntakeId"]
        .nunique()
        .rename("sessionsWithFieldRecord")
        .reset_index()
        .rename(columns={"fieldId": "questionId"})
    )

    question_summary_df = question_summary_df.merge(
        field_session_counts,
        on="questionId",
        how="outer",
    )

    if "firstChangedAt" in fields_df.columns:
        valid_change_rows = fields_df[
            fields_df["firstChangedAt"].notna()
        ].copy()

        first_change_counts = (
            valid_change_rows.groupby("fieldId")
            .agg(
                sessionsWithFirstChange=(
                    "shortIntakeId",
                    "nunique",
                )
            )
            .reset_index()
            .rename(columns={"fieldId": "questionId"})
        )

        question_summary_df = question_summary_df.merge(
            first_change_counts,
            on="questionId",
            how="left",
        )

    count_columns = [
        "sessionsReached",
        "totalRecordedEvents",
        "repeatVisits",
        "sessionsEndingHere",
        "sessionsWithFieldRecord",
        "sessionsWithFirstChange",
    ]

    for column_name in count_columns:
        if column_name in question_summary_df.columns:
            question_summary_df[column_name] = (
                question_summary_df[column_name]
                .fillna(0)
                .astype(int)
            )

    for column_name in [
        "averageEventOrder",
        "medianEventOrder",
    ]:
        if column_name in question_summary_df.columns:
            question_summary_df[column_name] = (
                question_summary_df[column_name]
                .round(2)
            )

    question_summary_df = question_summary_df.sort_values(
        by=[
            "averageEventOrder",
            "questionId",
        ],
        na_position="last",
    )

    # --------------------------------------------------
    # Build high-level Summary Metrics
    # --------------------------------------------------
    valid_completion_times = session_summary_df[
        "completionMinutes"
    ].dropna()

    total_sessions = len(session_summary_df)

    completed_sessions = int(
        (
            session_summary_df["sessionStatus"]
            == "Completed"
        ).sum()
    )

    incomplete_sessions = int(
        (
            session_summary_df["sessionStatus"]
            == "Incomplete"
        ).sum()
    )

    completion_rate = (
        completed_sessions / total_sessions
        if total_sessions
        else 0
    )

    average_completion = (
        valid_completion_times.mean()
        if not valid_completion_times.empty
        else None
    )

    median_completion = (
        valid_completion_times.median()
        if not valid_completion_times.empty
        else None
    )

    fastest_completion = (
        valid_completion_times.min()
        if not valid_completion_times.empty
        else None
    )

    slowest_completion = (
        valid_completion_times.max()
        if not valid_completion_times.empty
        else None
    )

    average_events = (
        session_summary_df["numberOfEvents"].mean()
        if total_sessions
        else None
    )

    average_fields = (
        session_summary_df["numberOfFields"].mean()
        if total_sessions
        else None
    )

    most_revisited_question = None

    if (
        not question_summary_df.empty
        and "repeatVisits" in question_summary_df.columns
        and question_summary_df["repeatVisits"].max() > 0
    ):
        most_revisited_row = question_summary_df.loc[
            question_summary_df["repeatVisits"].idxmax()
        ]

        most_revisited_question = (
            f"{most_revisited_row['questionId']} "
            f"({int(most_revisited_row['repeatVisits'])} "
            f"repeat visits)"
        )

    most_common_final_question = None

    if (
        not question_summary_df.empty
        and "sessionsEndingHere"
        in question_summary_df.columns
        and question_summary_df["sessionsEndingHere"].max() > 0
    ):
        most_common_final_row = question_summary_df.loc[
            question_summary_df[
                "sessionsEndingHere"
            ].idxmax()
        ]

        most_common_final_question = (
            f"{most_common_final_row['questionId']} "
            f"({int(most_common_final_row['sessionsEndingHere'])} "
            f"sessions)"
        )

    summary_metrics_df = pd.DataFrame(
        [
            {
                "metric": "Total sessions",
                "value": total_sessions,
                "definition": (
                    "Number of rows in the Sessions sheet."
                ),
            },
            {
                "metric": "Completed sessions",
                "value": completed_sessions,
                "definition": (
                    "Sessions with a submittedAt timestamp."
                ),
            },
            {
                "metric": "Incomplete sessions",
                "value": incomplete_sessions,
                "definition": (
                    "Sessions without a submittedAt timestamp."
                ),
            },
            {
                "metric": "Completion rate",
                "value": round(completion_rate, 4),
                "definition": (
                    "Completed sessions divided by total sessions."
                ),
            },
            {
                "metric": "Average completion time (minutes)",
                "value": (
                    round(average_completion, 2)
                    if pd.notna(average_completion)
                    else None
                ),
                "definition": (
                    "Mean submittedAt minus startedAt."
                ),
            },
            {
                "metric": "Median completion time (minutes)",
                "value": (
                    round(median_completion, 2)
                    if pd.notna(median_completion)
                    else None
                ),
                "definition": (
                    "Median submittedAt minus startedAt."
                ),
            },
            {
                "metric": "Fastest completion time (minutes)",
                "value": (
                    round(fastest_completion, 2)
                    if pd.notna(fastest_completion)
                    else None
                ),
                "definition": (
                    "Smallest valid session completion time."
                ),
            },
            {
                "metric": "Slowest completion time (minutes)",
                "value": (
                    round(slowest_completion, 2)
                    if pd.notna(slowest_completion)
                    else None
                ),
                "definition": (
                    "Largest valid session completion time."
                ),
            },
            {
                "metric": "Average recorded events per session",
                "value": (
                    round(average_events, 2)
                    if pd.notna(average_events)
                    else None
                ),
                "definition": (
                    "Average number of rows in Events per session."
                ),
            },
            {
                "metric": "Average recorded fields per session",
                "value": (
                    round(average_fields, 2)
                    if pd.notna(average_fields)
                    else None
                ),
                "definition": (
                    "Average number of unique field records per session."
                ),
            },
            {
                "metric": "Most revisited question",
                "value": most_revisited_question,
                "definition": (
                    "Question with the most repeated event records "
                    "after its first appearance in each session."
                ),
            },
            {
                "metric": "Most common final recorded question",
                "value": most_common_final_question,
                "definition": (
                    "Question appearing most often as the last "
                    "recorded event in a session."
                ),
            },
            {
                "metric": "Existing data-quality issues",
                "value": len(quality_df),
                "definition": (
                    "Number of rows in the parsed Data Quality sheet."
                ),
            },
        ]
    )

    # --------------------------------------------------
    # Write workbook
    # --------------------------------------------------
    print("Writing UX metrics workbook...")

    session_summary_df = make_excel_safe(
        session_summary_df
    )

    question_summary_df = make_excel_safe(
        question_summary_df
    )

    session_paths_df = make_excel_safe(
        session_paths_df
    )

    summary_metrics_df = make_excel_safe(
        summary_metrics_df
    )

    OUTPUT_FILE.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl",
    ) as writer:
        summary_metrics_df.to_excel(
            writer,
            sheet_name="Summary Metrics",
            index=False,
        )

        session_summary_df.to_excel(
            writer,
            sheet_name="Session Summary",
            index=False,
        )

        question_summary_df.to_excel(
            writer,
            sheet_name="Question Summary",
            index=False,
        )

        session_paths_df.to_excel(
            writer,
            sheet_name="Session Paths",
            index=False,
        )

        auto_fit_columns(writer)

    print("UX metrics workbook complete.")
    print(f"Sessions analyzed: {len(session_summary_df)}")
    print(f"Questions analyzed: {len(question_summary_df)}")
    print(f"Saved to: {OUTPUT_FILE.resolve()}")


if __name__ == "__main__":
    main()