from pathlib import Path

import pandas as pd
from azure.data.tables import TableServiceClient
from azure.identity import DefaultAzureCredential


$AccountUrl = $env:AZURE_STORAGE_ACCOUNT_URL
TABLE_NAME = "FormAnswerTiming"

PROJECT_ROOT = Path(__file__).resolve().parent

OUTPUT_FILE = (
    PROJECT_ROOT
    / "data"
    / "raw"
    / "master_raw.xlsx"
)


def main() -> None:
    print("Connecting to Azure...")

    credential = DefaultAzureCredential()

    service_client = TableServiceClient(
        endpoint=ACCOUNT_URL,
        credential=credential,
    )

    table_client = service_client.get_table_client(TABLE_NAME)

    print("Querying all records...")

    entities = list(table_client.list_entities())

    if not entities:
        print("No records were found.")
        return

    rows = [dict(entity) for entity in entities]

    dataframe = pd.DataFrame(rows)

    OUTPUT_FILE.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    dataframe.to_excel(
        OUTPUT_FILE,
        index=False,
    )

    print(f"Exported {len(dataframe)} records.")
    print(f"Saved to: {OUTPUT_FILE.resolve()}")


if __name__ == "__main__":
    main()