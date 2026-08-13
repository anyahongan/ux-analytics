from azure.data.tables import TableServiceClient
from azure.identity import DefaultAzureCredential

$AccountUrl = $env:AZURE_STORAGE_ACCOUNT_URL
TABLE_NAME = "FormAnswerTiming"

credential = DefaultAzureCredential()

service_client = TableServiceClient(
    endpoint=ACCOUNT_URL,
    credential=credential,
)

table_client = service_client.get_table_client(TABLE_NAME)

print("Requesting one record from Azure...")

entities = table_client.list_entities(results_per_page=1)

for entity in entities:
    print(dict(entity))
    break

print("Finished.")