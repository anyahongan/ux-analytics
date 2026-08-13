from pathlib import Path
import json

import pandas as pd
from openpyxl.utils import get_column_letter


PROJECT_ROOT = Path(__file__).resolve().parent

INPUT_FILE = (
    PROJECT_ROOT
    / "data"
    / "raw"
    / "master_raw.xlsx"
)

OUTPUT_FILE = (
    PROJECT_ROOT
    / "data"
    / "parsed"
    / "master_parsed.xlsx"
)



def parse_json_value(value, expected_type):
    """
    Convert a JSON string into a Python list or dictionary.

    Returns:
        parsed_value: The converted JSON value
        error_message: None when successful, otherwise an explanation
    """
    if pd.isna(value) or value == "":
        return expected_type(), None

    if isinstance(value, expected_type):
        return value, None

    try:
        parsed_value = json.loads(value)

        if not isinstance(parsed_value, expected_type):
            error_message = (
                f"Expected {expected_type.__name__}, "
                f"but found {type(parsed_value).__name__}"
            )
            return expected_type(), error_message

        return parsed_value, None

    except (json.JSONDecodeError, TypeError) as error:
        return expected_type(), str(error)


def shorten_at_first_dash(value):
    """
    Keep only the text before the first dash.
    """
    if pd.isna(value):
        return None

    return str(value).split("-", 1)[0]


def shorten_at_first_underscore(value):
    """
    Keep only the text before the first underscore.
    """
    if pd.isna(value):
        return None

    return str(value).split("_", 1)[0]


def convert_utc_column_for_excel(
    dataframe,
    column_name,
):
    """
    Parse a timestamp column as UTC and remove timezone
    metadata so Excel can store it. The clock time remains UTC.
    """
    if column_name not in dataframe.columns:
        return

    dataframe[column_name] = (
        pd.to_datetime(
            dataframe[column_name],
            errors="coerce",
            utc=True,
        )
        .dt.tz_localize(None)
    )


def auto_fit_excel_columns(writer):
    """
    Automatically widen Excel columns based on their contents.
    """
    for worksheet in writer.book.worksheets:
        for column_cells in worksheet.columns:
            maximum_length = 0
            column_number = column_cells[0].column
            column_letter = get_column_letter(column_number)

            for cell in column_cells:
                if cell.value is None:
                    continue

                cell_length = len(str(cell.value))
                maximum_length = max(
                    maximum_length,
                    cell_length,
                )

            adjusted_width = min(
                maximum_length + 2,
                50,
            )

            worksheet.column_dimensions[
                column_letter
            ].width = adjusted_width


def main():
    print("Reading raw Excel file...")

    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Could not find: {INPUT_FILE.resolve()}"
        )

    raw_df = pd.read_excel(INPUT_FILE)

    sessions_rows = []
    events_rows = []
    fields_rows = []
    quality_rows = []

    for row_number, row in raw_df.iterrows():
        source_row = row_number + 2

        intake_id = row.get("intakeId")
        row_key = row.get("RowKey")
        partition_key = row.get("PartitionKey")

        short_intake_id = shorten_at_first_dash(
            intake_id
        )

        short_row_key = shorten_at_first_underscore(
            row_key
        )

        sessions_rows.append(
            {
                "shortIntakeId": short_intake_id,
                "shortRowKey": short_row_key,
                "formType": row.get("formType"),
                "startedAt": row.get("startedAt"),
                "submittedAt": row.get("submittedAt"),
                "submissionType": row.get(
                    "submissionType"
                ),
                "isEvent": row.get("isEvent"),
                "intakeId": intake_id,
                "RowKey": row_key,
                "PartitionKey": partition_key,
                "Timestamp": row.get("Timestamp"),
                "sourceRow": source_row,
            }
        )

        events, events_error = parse_json_value(
            row.get("eventsJson"),
            list,
        )

        if events_error:
            quality_rows.append(
                {
                    "sourceRow": source_row,
                    "shortIntakeId": short_intake_id,
                    "intakeId": intake_id,
                    "field": "eventsJson",
                    "issueType": "Invalid JSON",
                    "details": events_error,
                }
            )

        for event_order, event in enumerate(
            events,
            start=1,
        ):
            if not isinstance(event, dict):
                quality_rows.append(
                    {
                        "sourceRow": source_row,
                        "shortIntakeId": short_intake_id,
                        "intakeId": intake_id,
                        "field": "eventsJson",
                        "issueType": (
                            "Unexpected event format"
                        ),
                        "details": str(event),
                    }
                )
                continue

            events_rows.append(
                {
                    "shortIntakeId": short_intake_id,
                    "shortRowKey": short_row_key,
                    "eventOrder": event_order,
                    "linkId": event.get("linkId"),
                    "eventTime": event.get("at"),
                    "RowKey": row_key,
                    "PartitionKey": partition_key,
                }
            )

        fields, fields_error = parse_json_value(
            row.get("fieldsJson"),
            dict,
        )

        if fields_error:
            quality_rows.append(
                {
                    "sourceRow": source_row,
                    "shortIntakeId": short_intake_id,
                    "intakeId": intake_id,
                    "field": "fieldsJson",
                    "issueType": "Invalid JSON",
                    "details": fields_error,
                }
            )

        for field_id, field_details in fields.items():
            if not isinstance(
                field_details,
                dict,
            ):
                quality_rows.append(
                    {
                        "sourceRow": source_row,
                        "shortIntakeId": short_intake_id,
                        "intakeId": intake_id,
                        "field": "fieldsJson",
                        "issueType": (
                            "Unexpected field format"
                        ),
                        "details": (
                            f"{field_id}: "
                            f"{field_details}"
                        ),
                    }
                )
                continue

            field_record = {
                "shortIntakeId": short_intake_id,
                "shortRowKey": short_row_key,
                "fieldId": field_id,
                "RowKey": row_key,
                "PartitionKey": partition_key,
            }

            for (
                property_name,
                property_value,
            ) in field_details.items():
                field_record[
                    property_name
                ] = property_value

            fields_rows.append(field_record)

    sessions_df = pd.DataFrame(sessions_rows)
    events_df = pd.DataFrame(events_rows)
    fields_df = pd.DataFrame(fields_rows)

    sessions_df["startedAt"] = pd.to_datetime(
        sessions_df["startedAt"],
        errors="coerce",
        utc=True,
    )

    sessions_df["submittedAt"] = pd.to_datetime(
        sessions_df["submittedAt"],
        errors="coerce",
        utc=True,
    )

    sessions_df["completionSeconds"] = (
        sessions_df["submittedAt"]
        - sessions_df["startedAt"]
    ).dt.total_seconds()

    sessions_df["completionMinutes"] = (
        sessions_df["completionSeconds"] / 60
    )

    for _, session in sessions_df.iterrows():
        if pd.isna(session["startedAt"]):
            quality_rows.append(
                {
                    "sourceRow": session.get(
                        "sourceRow"
                    ),
                    "shortIntakeId": session.get(
                        "shortIntakeId"
                    ),
                    "intakeId": session.get(
                        "intakeId"
                    ),
                    "field": "startedAt",
                    "issueType": (
                        "Missing or invalid timestamp"
                    ),
                    "details": (
                        "Could not identify a valid "
                        "session start time."
                    ),
                }
            )

        if pd.isna(session["submittedAt"]):
            quality_rows.append(
                {
                    "sourceRow": session.get(
                        "sourceRow"
                    ),
                    "shortIntakeId": session.get(
                        "shortIntakeId"
                    ),
                    "intakeId": session.get(
                        "intakeId"
                    ),
                    "field": "submittedAt",
                    "issueType": (
                        "Missing or invalid timestamp"
                    ),
                    "details": (
                        "Session may be incomplete."
                    ),
                }
            )

        completion_seconds = session.get(
            "completionSeconds"
        )

        if (
            pd.notna(completion_seconds)
            and completion_seconds < 0
        ):
            quality_rows.append(
                {
                    "sourceRow": session.get(
                        "sourceRow"
                    ),
                    "shortIntakeId": session.get(
                        "shortIntakeId"
                    ),
                    "intakeId": session.get(
                        "intakeId"
                    ),
                    "field": "completionSeconds",
                    "issueType": "Negative duration",
                    "details": completion_seconds,
                }
            )

    duplicate_intakes = sessions_df[
        sessions_df["intakeId"].notna()
        & sessions_df["intakeId"].duplicated(
            keep=False
        )
    ]

    for _, session in duplicate_intakes.iterrows():
        quality_rows.append(
            {
                "sourceRow": session.get(
                    "sourceRow"
                ),
                "shortIntakeId": session.get(
                    "shortIntakeId"
                ),
                "intakeId": session.get(
                    "intakeId"
                ),
                "field": "intakeId",
                "issueType": "Duplicate intake ID",
                "details": (
                    "This full intakeId appears "
                    "more than once."
                ),
            }
        )

    sessions_df["startedAt"] = (
        sessions_df["startedAt"]
        .dt.tz_localize(None)
    )

    sessions_df["submittedAt"] = (
        sessions_df["submittedAt"]
        .dt.tz_localize(None)
    )

    convert_utc_column_for_excel(
        sessions_df,
        "Timestamp",
    )

    convert_utc_column_for_excel(
        events_df,
        "eventTime",
    )

    field_timestamp_columns = [
        "firstChangedAt",
        "lastChangedAt",
        "changedAt",
        "createdAt",
        "updatedAt",
    ]

    for column_name in field_timestamp_columns:
        convert_utc_column_for_excel(
            fields_df,
            column_name,
        )

    sessions_df["completionSeconds"] = (
        sessions_df["completionSeconds"].round(2)
    )

    sessions_df["completionMinutes"] = (
        sessions_df["completionMinutes"].round(2)
    )

    sessions_df = sessions_df[
        [
            "shortIntakeId",
            "shortRowKey",
            "formType",
            "startedAt",
            "submittedAt",
            "completionSeconds",
            "completionMinutes",
            "submissionType",
            "isEvent",
            "intakeId",
            "RowKey",
            "PartitionKey",
            "Timestamp",
            "sourceRow",
        ]
    ]

    quality_df = pd.DataFrame(
        quality_rows,
        columns=[
            "sourceRow",
            "shortIntakeId",
            "intakeId",
            "field",
            "issueType",
            "details",
        ],
    )

    OUTPUT_FILE.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    print("Writing parsed workbook...")

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl",
    ) as writer:
        raw_df.to_excel(
            writer,
            sheet_name="Raw Data",
            index=False,
        )

        sessions_df.to_excel(
            writer,
            sheet_name="Sessions",
            index=False,
        )

        events_df.to_excel(
            writer,
            sheet_name="Events",
            index=False,
        )

        fields_df.to_excel(
            writer,
            sheet_name="Fields",
            index=False,
        )

        quality_df.to_excel(
            writer,
            sheet_name="Data Quality",
            index=False,
        )

        auto_fit_excel_columns(writer)

    print("Parsing complete.")
    print(f"Sessions: {len(sessions_df)}")
    print(f"Events: {len(events_df)}")
    print(f"Fields: {len(fields_df)}")
    print(
        f"Data quality issues: "
        f"{len(quality_df)}"
    )
    print("Timezone: UTC")
    print(f"Saved to: {OUTPUT_FILE.resolve()}")


if __name__ == "__main__":
    main()