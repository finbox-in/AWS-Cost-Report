# AWS-Cost-Report
Script to generate a Cost Report in XLSX format, with different sheets indicating different actionable reports

## Sheets with config.json keys
1. **Expensive Services**: It lists the expensive service names on your AWS Account. This requires Cost Explorer to be enabled on your account.
    Its config.json key looks like this:
    ```json
    "expensive_services": {
            "enabled": true,
            "past_days": 2,
            "cost_percentage": 80
        }
    ```
    
    | Key | Type | Description |
    | --- | ---  | --- |
    | `enabled` | Boolean | This specifies whether to include the sheet in the report workbook |
    | `past_days` | Integer | Looks for cost data for specified number of past days |
    | `cost_percentage` | Integer | Lists only the services for which sum of costs is `cost_percentage`% of total cost over the past `past_days` |

2. **Untagged Resources**: It lists all the resources for which specified tags are missing. Proper tagging always helps in analyzing the costs well over the cost explorer. It currently looks for the tags in Lambda Functions, DynamoDB Tables, EC2 Instances, Kinesis Streams, Firehose Delivery Streams and S3 Buckets.
    Its config.json key looks like this:
    ```json
    "untagged_resources": {
        "enabled": true,
        "tags": ["Name", "STAGE", "Pipeline"]
    }
    ```
    | Key | Type | Description |
    | --- | --- | --- |
    | `enabled` | Boolean | This specifies whether to include the sheet in the report workbook |
    | `tags` | List of Strings | Specifies the tags to look for in resources |
3. **Unreferenced Snapshots**: It lists all the unreferenced snapshots, i.e. the ones who have at least one of the volume, AMI or instance not referenced.
    Its config.json key looks like this:
    ```json
    "unreferenced_snapshots": {
        "enabled": true
    }
    ```
    | Key | Type | Description |
    | --- | --- | --- |
    | `enabled` | Boolean | This specifies whether to include the sheet in the report workbook |
4. **Unattached Volumes**: It lists all the volumes unattached to any EC2 Instance.
    Its config.json key looks like this:
    ```json
    "unattached_volumes": {
        "enabled": true
    }
    ```
    | Key | Type | Description |
    | --- | --- | --- |
    | `enabled` | Boolean | This specifies whether to include the sheet in the report workbook |
5. **Expensive Lambdas**: It lists the expensive Lambda Functions on your AWS Account. This requires Cost Explorer to be enabled on your account.
    Its config.json key looks like this:
    ```json
    "expensive_lambda_functions": {
        "enabled": true,
        "name_tag_key": "Name",
        "cost_percentage": 80,
        "past_days": 7
    }
    ```
    | Key | Type | Description |
    | --- | --- | --- |
    | `enabled` | Boolean | This specifies whether to include the sheet in the report workbook |
    | `name_tag_key` | String | This specifies which tag in the lambda function specifies the name of the function |
    | `past_days` | Integer | Looks for cost data for specified number of past days |
    | `cost_percentage` | Integer | Lists only the lambdas for which sum of costs is `cost_percentage`% of total lambda cost over the past `past_days` |
6. **Expensive Kinesis Streams**: It lists the expensive Kinesis Streams on your AWS Account. This requires Cost Explorer to be enabled on your account.
    Its config.json key looks like this:
    ```json
    "expensive_kinesis_streams": {
        "enabled": true,
        "name_tag_key": "Name",
        "cost_percentage": 80,
        "past_days": 7
    }
    ```
    | Key | Type | Description |
    | --- | --- | --- |
    | `enabled` | Boolean | This specifies whether to include the sheet in the report workbook |
    | `name_tag_key` | String | This specifies which tag in the kinesis stream specifies the name of the stream |
    | `past_days` | Integer | Looks for cost data for specified number of past days |
    | `cost_percentage` | Integer | Lists only the streams for which sum of costs is `cost_percentage`% of total Kinesis cost over the past `past_days` |
7. **Top N CloudWatch Log Groups by incoming bytes**: It lists the top N CloudWatch Log Groups with highest incoming bytes on your AWS Account over the specified past days. This requires Cost Explorer to be enabled on your account.
    Its config.json key looks like this:
    ```json
    "storage_cloudwatch_log_groups": {
        "enabled": true,
        "top_n": 10,
        "past_days": 14
    }
    ```
    | Key | Type | Description |
    | --- | --- | --- |
    | `enabled` | Boolean | This specifies whether to include the sheet in the report workbook |
    | `past_days` | Integer | Looks for cost data for specified number of past days |
    | `top_n` | Integer | Specifies the value of N |
8. **Unused Elastic IPs**: Lists the unused elastic IPs, the ones unassociated as well as the ones for which the instance is stopped.
    Its config.json key looks like this:
    ```json
    "unused_elastic_ips": {
        "enabled": true
    }
    ```
    | Key | Type | Description |
    | --- | --- | --- |
    | `enabled` | Boolean | This specifies whether to include the sheet in the report workbook |
