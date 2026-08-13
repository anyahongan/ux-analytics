from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parent

PARSED_FILE = (
    PROJECT_ROOT
    / "data"
    / "parsed"
    / "master_parsed.xlsx"
)

CURRENT_REPORTS_FOLDER = (
    PROJECT_ROOT
    / "reports"
    / "current"
)

ARCHIVE_REPORTS_FOLDER = (
    PROJECT_ROOT
    / "reports"
    / "archive"
)

RUN_METADATA_FILE = (
    PROJECT_ROOT
    / "config"
    / "current_run.json"
)


PIPELINE_STEPS = [
    {
        "name": "Export raw Azure data",
        "script": PROJECT_ROOT / "export_raw_data.py",
        "expected_output": (
            PROJECT_ROOT
            / "data"
            / "raw"
            / "master_raw.xlsx"
        ),
    },
    {
        "name": "Parse and standardize UX data",
        "script": PROJECT_ROOT / "parse_ux_data.py",
        "expected_output": PARSED_FILE,
    },
    {
        "name": "Build UX metrics",
        "script": PROJECT_ROOT / "build_ux_metrics.py",
        "expected_output": (
            PROJECT_ROOT
            / "data"
            / "parsed"
            / "ux_metrics.xlsx"
        ),
    },
    {
        "name": "Build UX report",
        "script": PROJECT_ROOT / "build_ux_diagnostics.py",
        "expected_output": (
            CURRENT_REPORTS_FOLDER
            / "ux_report.xlsx"
        ),
    },
    {
        "name": "Validate UX pipeline",
        "script": PROJECT_ROOT / "validate_ux_pipeline.py",
        "expected_output": (
            CURRENT_REPORTS_FOLDER
            / "ux_validation_report.xlsx"
        ),
    },
]


def parse_arguments() -> argparse.Namespace:
    """
    Read optional UTC date or date-time filters.

    Examples:

    All available data:
        python complete_ux_metrics.py

    Whole-date range:
        python complete_ux_metrics.py \
            --start-date 2026-07-01 \
            --end-date 2026-07-31

    Time range:
        python complete_ux_metrics.py \
            --start-time "2026-07-24 09:00" \
            --end-time "2026-07-24 12:00"
    """
    parser = argparse.ArgumentParser(
        description=(
            "Run the complete UX analytics pipeline."
        )
    )

    parser.add_argument(
        "--start-date",
        type=str,
        default=None,
        help=(
            "First date to include, formatted YYYY-MM-DD."
        ),
    )

    parser.add_argument(
        "--end-date",
        type=str,
        default=None,
        help=(
            "Last date to include, formatted YYYY-MM-DD. "
            "The full end date is included."
        ),
    )

    parser.add_argument(
        "--start-time",
        type=str,
        default=None,
        help=(
            "First UTC date-time to include, formatted "
            "\"YYYY-MM-DD HH:MM\"."
        ),
    )

    parser.add_argument(
        "--end-time",
        type=str,
        default=None,
        help=(
            "Last UTC date-time to include, formatted "
            "\"YYYY-MM-DD HH:MM\"."
        ),
    )

    arguments = parser.parse_args()

    date_arguments_used = (
        arguments.start_date is not None
        or arguments.end_date is not None
    )

    time_arguments_used = (
        arguments.start_time is not None
        or arguments.end_time is not None
    )

    if date_arguments_used and time_arguments_used:
        parser.error(
            "Use either --start-date/--end-date or "
            "--start-time/--end-time, not both."
        )

    if (
        arguments.start_date is None
        and arguments.end_date is not None
    ):
        parser.error(
            "--end-date requires --start-date."
        )

    if (
        arguments.start_date is not None
        and arguments.end_date is None
    ):
        parser.error(
            "--start-date requires --end-date."
        )

    if (
        arguments.start_time is None
        and arguments.end_time is not None
    ):
        parser.error(
            "--end-time requires --start-time."
        )

    if (
        arguments.start_time is not None
        and arguments.end_time is None
    ):
        parser.error(
            "--start-time requires --end-time."
        )

    return arguments


def parse_exact_datetime(
    value: str,
    argument_name: str,
    expected_format: str,
) -> pd.Timestamp:
    """
    Parse a value using one exact date or date-time format.
    """
    try:
        parsed_value = datetime.strptime(
            value,
            expected_format,
        )
    except ValueError as error:
        readable_format = (
            "YYYY-MM-DD"
            if expected_format == "%Y-%m-%d"
            else "YYYY-MM-DD HH:MM"
        )

        raise ValueError(
            f"{argument_name} must use "
            f"{readable_format}. Received: {value}"
        ) from error

    return pd.Timestamp(parsed_value)


def resolve_analysis_window(
    arguments: argparse.Namespace,
) -> tuple[
    pd.Timestamp | None,
    pd.Timestamp | None,
    str,
]:
    """
    Resolve the selected UTC analysis window.

    Returns:
        start_boundary
        end_boundary
        filter_mode

    filter_mode is one of:
        all
        date
        time
    """
    if (
        arguments.start_date is None
        and arguments.start_time is None
    ):
        return None, None, "all"

    if arguments.start_date is not None:
        start_boundary = parse_exact_datetime(
            arguments.start_date,
            "--start-date",
            "%Y-%m-%d",
        )

        inclusive_end_date = parse_exact_datetime(
            arguments.end_date,
            "--end-date",
            "%Y-%m-%d",
        )

        if inclusive_end_date < start_boundary:
            raise ValueError(
                "--end-date cannot occur before "
                "--start-date."
            )

        exclusive_end_boundary = (
            inclusive_end_date
            + pd.Timedelta(days=1)
        )

        return (
            start_boundary,
            exclusive_end_boundary,
            "date",
        )

    start_boundary = parse_exact_datetime(
        arguments.start_time,
        "--start-time",
        "%Y-%m-%d %H:%M",
    )

    inclusive_end_boundary = parse_exact_datetime(
        arguments.end_time,
        "--end-time",
        "%Y-%m-%d %H:%M",
    )

    if inclusive_end_boundary < start_boundary:
        raise ValueError(
            "--end-time cannot occur before "
            "--start-time."
        )

    return (
        start_boundary,
        inclusive_end_boundary,
        "time",
    )


def create_required_folders() -> None:
    """
    Create all folders required by the pipeline.
    """
    required_folders = [
        PROJECT_ROOT / "data" / "raw",
        PROJECT_ROOT / "data" / "parsed",
        CURRENT_REPORTS_FOLDER,
        ARCHIVE_REPORTS_FOLDER,
        PROJECT_ROOT / "config",
    ]

    for folder in required_folders:
        folder.mkdir(
            parents=True,
            exist_ok=True,
        )


def run_script(
    name: str,
    script_path: Path,
    expected_output: Path,
) -> None:
    """
    Run one pipeline script and confirm its output.
    """
    if not script_path.exists():
        raise FileNotFoundError(
            f"Could not find the required script:\n"
            f"{script_path}"
        )

    print()
    print("=" * 70)
    print(f"Starting: {name}")
    print(f"Script: {script_path.name}")
    print("=" * 70)

    step_start = time.perf_counter()

    result = subprocess.run(
        [
            sys.executable,
            str(script_path),
        ],
        cwd=PROJECT_ROOT,
        check=False,
    )

    elapsed_seconds = (
        time.perf_counter()
        - step_start
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"{name} failed after "
            f"{elapsed_seconds:.1f} seconds "
            f"with exit code {result.returncode}."
        )

    if not expected_output.exists():
        raise FileNotFoundError(
            f"{name} finished, but the expected output "
            f"was not created:\n"
            f"{expected_output}"
        )

    print()
    print(
        f"Completed: {name} "
        f"({elapsed_seconds:.1f} seconds)"
    )
    print(f"Confirmed output: {expected_output}")


def read_parsed_workbook() -> dict[str, pd.DataFrame]:
    """
    Read every worksheet from the parsed workbook.
    """
    if not PARSED_FILE.exists():
        raise FileNotFoundError(
            f"Could not find parsed workbook:\n"
            f"{PARSED_FILE}"
        )

    return pd.read_excel(
        PARSED_FILE,
        sheet_name=None,
    )


def filter_parsed_workbook(
    start_boundary: pd.Timestamp | None,
    end_boundary: pd.Timestamp | None,
    filter_mode: str,
) -> int:
    """
    Filter Sessions by UTC startedAt time.

    Events, Fields, and Data Quality rows are retained only
    when they belong to an included session.
    """
    workbook_sheets = read_parsed_workbook()

    required_sheets = {
        "Sessions",
        "Events",
        "Fields",
        "Data Quality",
    }

    missing_sheets = (
        required_sheets
        - set(workbook_sheets)
    )

    if missing_sheets:
        raise ValueError(
            "The parsed workbook is missing these sheets: "
            f"{sorted(missing_sheets)}"
        )

    sessions_df = workbook_sheets["Sessions"].copy()
    events_df = workbook_sheets["Events"].copy()
    fields_df = workbook_sheets["Fields"].copy()
    quality_df = workbook_sheets["Data Quality"].copy()

    if "shortIntakeId" not in sessions_df.columns:
        raise ValueError(
            "The Sessions sheet is missing shortIntakeId."
        )

    if "startedAt" not in sessions_df.columns:
        raise ValueError(
            "The Sessions sheet is missing startedAt."
        )

    sessions_df["startedAt"] = pd.to_datetime(
        sessions_df["startedAt"],
        errors="coerce",
    )

    if "submittedAt" in sessions_df.columns:
        sessions_df["submittedAt"] = pd.to_datetime(
            sessions_df["submittedAt"],
            errors="coerce",
        )

    if (
        start_boundary is not None
        and end_boundary is not None
    ):
        if filter_mode == "date":
            date_filter = (
                sessions_df["startedAt"].notna()
                & (
                    sessions_df["startedAt"]
                    >= start_boundary
                )
                & (
                    sessions_df["startedAt"]
                    < end_boundary
                )
            )

        else:
            date_filter = (
                sessions_df["startedAt"].notna()
                & (
                    sessions_df["startedAt"]
                    >= start_boundary
                )
                & (
                    sessions_df["startedAt"]
                    <= end_boundary
                )
            )

        sessions_df = sessions_df[
            date_filter
        ].copy()

    included_session_ids = set(
        sessions_df["shortIntakeId"]
        .dropna()
        .astype(str)
    )

    if "shortIntakeId" in events_df.columns:
        events_df = events_df[
            events_df["shortIntakeId"]
            .astype(str)
            .isin(included_session_ids)
        ].copy()

    if "shortIntakeId" in fields_df.columns:
        fields_df = fields_df[
            fields_df["shortIntakeId"]
            .astype(str)
            .isin(included_session_ids)
        ].copy()

    if (
        not quality_df.empty
        and "shortIntakeId"
        in quality_df.columns
    ):
        quality_df = quality_df[
            quality_df["shortIntakeId"]
            .astype(str)
            .isin(included_session_ids)
        ].copy()

    if sessions_df.empty:
        raise ValueError(
            "No sessions were found in the selected "
            "UTC window."
        )

    workbook_sheets["Sessions"] = sessions_df
    workbook_sheets["Events"] = events_df
    workbook_sheets["Fields"] = fields_df
    workbook_sheets["Data Quality"] = quality_df

    with pd.ExcelWriter(
        PARSED_FILE,
        engine="openpyxl",
    ) as writer:
        for sheet_name, dataframe in workbook_sheets.items():
            dataframe.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
            )

    return len(sessions_df)


def load_previous_run_metadata() -> dict:
    """
    Read metadata for the reports currently in current.
    """
    if not RUN_METADATA_FILE.exists():
        return {}

    try:
        with RUN_METADATA_FILE.open(
            "r",
            encoding="utf-8",
        ) as metadata_file:
            return json.load(metadata_file)

    except (
        json.JSONDecodeError,
        OSError,
    ):
        return {}


def sanitize_archive_value(value: str) -> str:
    """
    Make a date-time safe for use in a filename.
    """
    return (
        value
        .replace(":", "-")
        .replace(" ", "_")
    )


def make_archive_label(metadata: dict) -> str:
    """
    Create a readable archive filename label.
    """
    start_value = metadata.get("start_value")
    end_value = metadata.get("end_value")
    analysis_scope = metadata.get(
        "analysis_scope"
    )
    run_timestamp = metadata.get(
        "run_timestamp"
    )

    if start_value and end_value:
        return (
            f"{sanitize_archive_value(start_value)}"
            f"_to_"
            f"{sanitize_archive_value(end_value)}"
            "_UTC"
        )

    if analysis_scope == "all_available_data":
        if run_timestamp:
            return f"all_data_{run_timestamp}_UTC"

        return "all_data_UTC"

    return datetime.now().strftime(
        "%Y-%m-%d_%H-%M-%S"
    )


def make_unique_archive_path(
    desired_path: Path,
) -> Path:
    """
    Prevent archived reports from overwriting each other.
    """
    if not desired_path.exists():
        return desired_path

    counter = 2

    while True:
        candidate_path = (
            desired_path.parent
            / (
                f"{desired_path.stem}"
                f"_{counter}"
                f"{desired_path.suffix}"
            )
        )

        if not candidate_path.exists():
            return candidate_path

        counter += 1


def archive_current_reports() -> list[
    tuple[Path, Path]
]:
    """
    Move existing current reports into archive.

    Returns pairs of:
        archived_path
        original_current_path
    """
    current_report_files = [
        CURRENT_REPORTS_FOLDER
        / "ux_report.xlsx",
        CURRENT_REPORTS_FOLDER
        / "ux_validation_report.xlsx",
    ]

    existing_report_files = [
        file_path
        for file_path in current_report_files
        if file_path.exists()
    ]

    if not existing_report_files:
        return []

    previous_metadata = (
        load_previous_run_metadata()
    )

    archive_label = make_archive_label(
        previous_metadata
    )

    archived_file_pairs = []

    print()
    print("Archiving previous current reports...")

    for current_file in existing_report_files:
        desired_archive_path = (
            ARCHIVE_REPORTS_FOLDER
            / (
                f"{current_file.stem}"
                f"_{archive_label}"
                f"{current_file.suffix}"
            )
        )

        archive_path = make_unique_archive_path(
            desired_archive_path
        )

        shutil.move(
            str(current_file),
            str(archive_path),
        )

        archived_file_pairs.append(
            (
                archive_path,
                current_file,
            )
        )

        print(
            f"Archived: {archive_path.name}"
        )

    return archived_file_pairs


def restore_archived_reports(
    archived_file_pairs: list[
        tuple[Path, Path]
    ],
) -> None:
    """
    Restore prior reports if new report generation fails.
    """
    if not archived_file_pairs:
        return

    print()
    print(
        "Restoring previous current reports because "
        "the new report run failed..."
    )

    for (
        archived_path,
        current_path,
    ) in archived_file_pairs:
        if archived_path.exists():
            shutil.move(
                str(archived_path),
                str(current_path),
            )


def save_run_metadata(
    arguments: argparse.Namespace,
    filter_mode: str,
    session_count: int,
) -> None:
    """
    Save information about reports currently in current.
    """
    run_timestamp = datetime.now().strftime(
        "%Y-%m-%d_%H-%M-%S"
    )

    if filter_mode == "all":
        metadata = {
            "analysis_scope": "all_available_data",
            "start_value": None,
            "end_value": None,
            "session_count": session_count,
            "run_timestamp": run_timestamp,
            "timezone": "UTC",
        }

    elif filter_mode == "date":
        metadata = {
            "analysis_scope": "date_range",
            "start_value": arguments.start_date,
            "end_value": arguments.end_date,
            "session_count": session_count,
            "run_timestamp": run_timestamp,
            "timezone": "UTC",
        }

    else:
        metadata = {
            "analysis_scope": "time_range",
            "start_value": arguments.start_time,
            "end_value": arguments.end_time,
            "session_count": session_count,
            "run_timestamp": run_timestamp,
            "timezone": "UTC",
        }

    with RUN_METADATA_FILE.open(
        "w",
        encoding="utf-8",
    ) as metadata_file:
        json.dump(
            metadata,
            metadata_file,
            indent=4,
        )


def print_analysis_period(
    arguments: argparse.Namespace,
    filter_mode: str,
) -> None:
    """
    Display the selected UTC analysis window.
    """
    print()

    if filter_mode == "all":
        print(
            "Analysis period: All available Azure data"
        )

    elif filter_mode == "date":
        print(
            "Analysis period: "
            f"{arguments.start_date} through "
            f"{arguments.end_date}"
        )
        print("Timezone: UTC")

    else:
        print(
            "Analysis period: "
            f"{arguments.start_time} through "
            f"{arguments.end_time}"
        )
        print("Timezone: UTC")


def main() -> None:
    """
    Run the complete UX analytics pipeline.
    """
    arguments = parse_arguments()

    (
        start_boundary,
        end_boundary,
        filter_mode,
    ) = resolve_analysis_window(arguments)

    pipeline_start = time.perf_counter()

    print()
    print("UX Analytics Pipeline")
    print("-----------------------------")
    print("This will:")
    print("1. Query Azure Table Storage")
    print("2. Update the raw master data")
    print("3. Parse and standardize the data")
    print("4. Apply the selected UTC window")
    print("5. Generate UX metrics")
    print("6. Archive the previous reports")
    print("7. Build the new UX report")
    print("8. Validate the analytics pipeline")

    print_analysis_period(
        arguments,
        filter_mode,
    )

    create_required_folders()

    archived_file_pairs = []

    try:
        run_script(
            name=PIPELINE_STEPS[0]["name"],
            script_path=PIPELINE_STEPS[0]["script"],
            expected_output=(
                PIPELINE_STEPS[0][
                    "expected_output"
                ]
            ),
        )

        run_script(
            name=PIPELINE_STEPS[1]["name"],
            script_path=PIPELINE_STEPS[1]["script"],
            expected_output=(
                PIPELINE_STEPS[1][
                    "expected_output"
                ]
            ),
        )

        print()
        print("=" * 70)
        print("Applying analysis window")
        print("=" * 70)

        included_session_count = (
            filter_parsed_workbook(
                start_boundary,
                end_boundary,
                filter_mode,
            )
        )

        print(
            f"Sessions included: "
            f"{included_session_count}"
        )

        print_analysis_period(
            arguments,
            filter_mode,
        )

        run_script(
            name=PIPELINE_STEPS[2]["name"],
            script_path=PIPELINE_STEPS[2]["script"],
            expected_output=(
                PIPELINE_STEPS[2][
                    "expected_output"
                ]
            ),
        )

        archived_file_pairs = (
            archive_current_reports()
        )

        run_script(
            name=PIPELINE_STEPS[3]["name"],
            script_path=PIPELINE_STEPS[3]["script"],
            expected_output=(
                PIPELINE_STEPS[3][
                    "expected_output"
                ]
            ),
        )

        run_script(
            name=PIPELINE_STEPS[4]["name"],
            script_path=PIPELINE_STEPS[4]["script"],
            expected_output=(
                PIPELINE_STEPS[4][
                    "expected_output"
                ]
            ),
        )

        save_run_metadata(
            arguments,
            filter_mode,
            included_session_count,
        )

    except Exception:
        restore_archived_reports(
            archived_file_pairs
        )
        raise

    total_seconds = (
        time.perf_counter()
        - pipeline_start
    )

    main_report = (
        CURRENT_REPORTS_FOLDER
        / "ux_report.xlsx"
    )

    validation_report = (
        CURRENT_REPORTS_FOLDER
        / "ux_validation_report.xlsx"
    )

    print()
    print("=" * 70)
    print("Pipeline completed successfully.")
    print(
        f"Total runtime: "
        f"{total_seconds:.1f} seconds"
    )
    print()
    print(
        f"Sessions included: "
        f"{included_session_count}"
    )

    print_analysis_period(
        arguments,
        filter_mode,
    )

    print()
    print("Main report:")
    print(main_report)
    print()
    print("Validation report:")
    print(validation_report)
    print("=" * 70)


if __name__ == "__main__":
    try:
        main()

    except KeyboardInterrupt:
        print()
        print("Pipeline stopped by user.")
        sys.exit(1)

    except Exception as error:
        print()
        print("=" * 70)
        print("Pipeline failed.")
        print(str(error))
        print("=" * 70)
        sys.exit(1)