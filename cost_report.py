"""
Generates report.xlsx file, based on the options specified on config.json file
"""

import boto3
import xlsxwriter
import json
from datetime import datetime, timedelta
from ebs_helpers import get_snapshots, get_available_volumes
from fetch_helpers import get_lambda_functions, get_dynamodb_tables
from fetch_helpers import get_ec2_reservations
from fetch_helpers import get_kinesis_streams, get_firehose_delivery_streams
from collections import defaultdict

# constants
OUTPUT_FILE_NAME = 'report.xlsx'
UNIQUE_ID_SIZE = 10
MAIN_HEADING_BG_COLOR = '#0080ff'
MAIN_HEADING_FONT_COLOR = '#ffffff'
MAIN_HEADING_FONT_SIZE = 13
SUB_HEADING_BG_COLOR = '#969696'
SUB_HEADING_FONT_COLOR = '#ffffff'
SUB_HEADING_FONT_SIZE = 13
CELL_FONT_SIZE = 13
GREEN_FONT_COLOR = '#037d50'
RED_FONT_COLOR = '#cc0000'

workbook = xlsxwriter.Workbook(OUTPUT_FILE_NAME)
main_heading = workbook.add_format({
    'font_color': MAIN_HEADING_FONT_COLOR,
    'bg_color': MAIN_HEADING_BG_COLOR,
    'valign': 'vcenter', 'border': 1,
    'font_size': MAIN_HEADING_FONT_SIZE
})
sub_heading = workbook.add_format({
    'font_color': SUB_HEADING_FONT_COLOR,
    'bg_color': SUB_HEADING_BG_COLOR,
    'valign': 'vcenter', 'border': 1,
    'font_size': SUB_HEADING_FONT_SIZE
})
generic_cell = workbook.add_format({
    'valign': 'vcenter', 'border': 1,
    'font_size': CELL_FONT_SIZE
})
green_text_cell = workbook.add_format({
    'valign': 'vcenter', 'border': 1,
    'font_size': CELL_FONT_SIZE,
    'font_color': GREEN_FONT_COLOR
})
red_text_cell = workbook.add_format({
    'valign': 'vcenter', 'border': 1,
    'font_size': CELL_FONT_SIZE,
    'font_color': RED_FONT_COLOR
})

print("Loading config.json file....")
config = dict()
with open("config.json") as json_file:
    config = json.load(json_file)
print("configuration loaded!")


def generate_unique_string():
    import random
    import string
    return "".join(random.choices(string.ascii_lowercase, k=UNIQUE_ID_SIZE))


def add_untagged_in_worksheet(worksheet, row, resource_type, resource_name, tags_to_look, tags):
    to_add = False
    values = [resource_type, resource_name]
    for key in tags_to_look:
        if (key not in tags) or (len(tags[key]) < 3):
            to_add = True
            values.append(False)
        else:
            values.append(True)
    if to_add:
        col = 0
        for value in values:
            if isinstance(value, bool):
                if value:
                    worksheet.write(row, col, "AVAILABLE", green_text_cell)
                else:
                    worksheet.write(row, col, "UNAVAILABLE", red_text_cell)
            else:
                worksheet.write(row, col, value, generic_cell)
            col += 1
        row += 1
    return row


def get_tags_dict_from_list(tags_list):
    tags = dict()
    for tag_dict in tag_list:
        tags[tag_dict['Key']] = tag_dict['Value']
    return tags


# conditional checks for all parameters
if config["expensive_services"]["enabled"]:
    cost_percentage = config["expensive_services"]["cost_percentage"]
    past_days = config["expensive_services"]["past_days"]
    # expensive services
    print("\nLooking for expensive services")
    print("---")
    start_date = (datetime.today() - timedelta(days=past_days)).strftime("%Y-%m-%d")
    end_date = datetime.today().strftime("%Y-%m-%d")

    services_worksheet = workbook.add_worksheet("{}% Cost Services".format(cost_percentage))
    row = 0
    # add headings
    services_worksheet.write(row, 0, "Service", main_heading)
    services_worksheet.write(row, 1, "Cost (in USD) for past {} days".format(past_days), main_heading)
    # set length of columns
    services_worksheet.set_column(0, 1, 60)
    row += 1
    # create cost explorer client
    client = boto3.client('ce')
    token = None
    results = []
    while True:
        if token:
            kwargs = {'NextPageToken': token}
        else:
            kwargs = {}
        data = client.get_cost_and_usage(TimePeriod={'Start': start_date, 'End': end_date}, Granularity='DAILY',
                                         Metrics=['UnblendedCost'], GroupBy=[{'Type': 'DIMENSION', 'Key': 'SERVICE'}],
                                         **kwargs)
        results += data['ResultsByTime']
        token = data.get('NextPageToken')
        if not token:
            break
    service_to_cost = defaultdict(int)
    total_cost = 0
    for result in results:
        for group in result.get("Groups", []):
            service_name = group["Keys"][0]
            cost = float(group["Metrics"]["UnblendedCost"]["Amount"])
            if cost > 0:
                service_to_cost[service_name] += cost
                total_cost += cost

    sorted_services = sorted(service_to_cost.items(), key=lambda x: x[1], reverse=True)

    target_amount = (cost_percentage / 100.0) * total_cost
    current_amount = 0
    for service_name, cost in sorted_services:
        services_worksheet.write(row, 0, service_name, generic_cell)
        services_worksheet.write(row, 1, cost, generic_cell)
        row += 1
        current_amount += cost
        if current_amount > target_amount:
            break
    if total_cost > 0:
        services_worksheet.write(row, 0, "ALL SERVICES", sub_heading)
        services_worksheet.write(row, 1, total_cost, sub_heading)
        row += 1


if config["untagged_resources"]["enabled"]:
    tags_to_look = config["untagged_resources"]["tags"]
    # untagged resources
    print("\nLooking for untagged resources")
    print("---")

    untagged_worksheet = workbook.add_worksheet("Untagged Resources")
    row = 0
    # add headings
    untagged_worksheet.write(row, 0, "Resource", main_heading)
    untagged_worksheet.write(row, 1, "Name", main_heading)
    col = 2
    for key in tags_to_look:
        untagged_worksheet.write(row, col, key, main_heading)
        col += 1
    row += 1
    # set length of columns
    untagged_worksheet.set_column(0, 0, 30)
    untagged_worksheet.set_column(1, 1, 70)
    untagged_worksheet.set_column(2, len(tags_to_look)+1, 14)

    # ------------------#
    # search in lambdas #
    # ------------------#
    print("\nFetching Lambda Functions...might take a while")
    client = boto3.client("lambda")
    # look for tags and add in worksheet if required
    for function in get_lambda_functions(client):
        tags = client.list_tags(Resource=function['FunctionArn']).get('Tags', dict())
        function_name = function['FunctionName']
        print("Checking for Lambda Function", function_name)
        row = add_untagged_in_worksheet(untagged_worksheet, row, "Lambda Function", function_name, tags_to_look, tags)

    # -------------------#
    # search in dynamodb #
    # -------------------#
    print("\nFetching DynamoDB Tables...might take a while")
    client = boto3.client("dynamodb")
    # look for tags and add in worksheet if required
    for table in get_dynamodb_tables(client):
        print("Checking for DynamoDB Table", table)
        table_arn = client.describe_table(TableName=table)['Table']['TableArn']
        tag_response = client.list_tags_of_resource(ResourceArn=table_arn)
        tag_list = tag_response.get('Tags', [])
        next_token = tag_response.get('NextToken')
        while next_token is not None:
            tag_response = client.list_tags_of_resource(ResourceArn=table_arn, NextToken=next_token)
            tag_list.extend(tag_response['Tags'])
            next_token = tag_response.get('NextToken')
        tags = get_tags_dict_from_list(tag_list)
        row = add_untagged_in_worksheet(untagged_worksheet, row, "DynamoDB Table", table, tags_to_look, tags)

    # ------------------#
    # search in EC2     #
    # ------------------#
    print("\nFetching EC2 Instances...might take a while")
    client = boto3.client("ec2")
    # look for tags and add in worksheet if required
    for reservation in get_ec2_reservations(client):
        for instance in reservation['Instances']:
            instance_identifier = instance['InstanceId']
            print("Checking for EC2 Instance", instance_identifier)
            tag_list = instance['Tags']
            tags = get_tags_dict_from_list(tag_list)
            if tags.get('Name') is not None:
                instance_identifier += (" (" + tags['Name'] + ")")
            row = add_untagged_in_worksheet(untagged_worksheet, row, "EC2 Instance", instance_identifier, tags_to_look, tags)

    # ------------------#
    # search in Kinesis #
    # ------------------#
    print("\nFetching Kinesis Streams...might take a while")
    client = boto3.client("kinesis")
    # look for tags and add in worksheet if required
    for stream in get_kinesis_streams(client):
        print("Checking for Kinesis Stream", stream)
        tag_response = client.list_tags_for_stream(StreamName=stream)
        tag_list = tag_response['Tags']
        has_more_tags = tag_response.get('HasMoreTags', False)
        while has_more_tags:
            tag_response = client.list_tags_for_stream(StreamName=stream, ExclusiveStartTagKey=tag_list[-1]['Key'])
            tag_list.extend(tag_response['Tags'])
            has_more_tags = tag_response.get('HasMoreTags', False)
        tags = get_tags_dict_from_list(tag_list)
        row = add_untagged_in_worksheet(untagged_worksheet, row, "Kinesis Stream", stream, tags_to_look, tags)

    # -------------------#
    # search in Firehose #
    # -------------------#
    print("\nFetching Firehose Delivery Streams...might take a while")
    client = boto3.client("firehose")
    # look for tags and add in worksheet if required
    for stream in get_firehose_delivery_streams(client):
        print("Checking for Firehose Delivery Streams", stream)
        tag_response = client.list_tags_for_delivery_stream(DeliveryStreamName=stream)
        tag_list = tag_response['Tags']
        has_more_tags = tag_response.get('HasMoreTags', False)
        while has_more_tags:
            tag_response = client.list_tags_for_delivery_stream(DeliveryStreamName=stream, ExclusiveStartTagKey=tag_list[-1]['Key'])
            tag_list.extend(tag_response['Tags'])
            has_more_tags = tag_response.get('HasMoreTags', False)
        tags = get_tags_dict_from_list(tag_list)
        row = add_untagged_in_worksheet(untagged_worksheet, row, "Firehose Delivery Stream", stream, tags_to_look, tags)

    # ------------------#
    #    search in S3   #
    # ------------------#
    print("\nFetching S3 Buckets...might take a while")
    client = boto3.client("s3")
    # fetch all S3 buckets and store in "buckets"
    buckets = client.list_buckets()['Buckets']
    # look for tags and add in worksheet if required
    for bucket in buckets:
        bucket_name = bucket['Name']
        print("Checking for S3 Bucket", bucket_name)
        try:
            tag_list = client.get_bucket_tagging(Bucket=bucket_name)['TagSet']
        except Exception as e:
            print(e)
            tag_list = []
        tags = get_tags_dict_from_list(tag_list)
        row = add_untagged_in_worksheet(untagged_worksheet, row, "S3 Bucket", bucket_name, tags_to_look, tags)

if config["unreferenced_snapshots"]["enabled"]:
    # Unreferenced Snapshots
    print("\nLooking for unreferenced snapshots")
    print("---")

    snapshots_worksheet = workbook.add_worksheet("Unreferenced Snapshots")
    row = 0
    # add headings
    snapshots_worksheet.write(row, 0, "Snapshot ID", main_heading)
    snapshots_worksheet.write(row, 1, "Size", main_heading)
    snapshots_worksheet.write(row, 2, "Start Time", main_heading)
    snapshots_worksheet.write(row, 3, "Volume", main_heading)
    snapshots_worksheet.write(row, 4, "AMI", main_heading)
    snapshots_worksheet.write(row, 5, "Instance", main_heading)
    snapshots_worksheet.write(row, 6, "Volume ID", main_heading)
    snapshots_worksheet.write(row, 7, "Volume Name", main_heading)
    snapshots_worksheet.write(row, 8, "AMI ID", main_heading)
    snapshots_worksheet.write(row, 9, "AMI Name", main_heading)
    snapshots_worksheet.write(row, 10, "Instance ID", main_heading)
    snapshots_worksheet.write(row, 11, "Instance Name", main_heading)
    # set length of columns
    snapshots_worksheet.set_column(0, 11, 30)

    row += 1
    print("Fetching snapshots...might take a while")
    for snapshot in get_snapshots():
        print("Checking for snapshot", snapshot['id'])
        if (not snapshot['volume_exists']) or (not snapshot['ami_exists']) or (not snapshot['instance_exists']):
            snapshots_worksheet.write(row, 0, snapshot['id'], generic_cell)
            snapshots_worksheet.write(row, 1, str(snapshot['size'])+" GB", generic_cell)
            snapshots_worksheet.write(row, 2, str(snapshot['start_time']), generic_cell)
            snapshots_worksheet.write(row, 3, snapshot['volume_exists'], green_text_cell if snapshot['volume_exists'] else red_text_cell)
            snapshots_worksheet.write(row, 4, snapshot['ami_exists'], green_text_cell if snapshot['ami_exists'] else red_text_cell)
            snapshots_worksheet.write(row, 5, snapshot['instance_exists'], green_text_cell if snapshot['instance_exists'] else red_text_cell)
            snapshots_worksheet.write(row, 6, snapshot['volume_id'], generic_cell)
            snapshots_worksheet.write(row, 7, snapshot['volume_name'], generic_cell)
            snapshots_worksheet.write(row, 8, snapshot['ami_id'], generic_cell)
            snapshots_worksheet.write(row, 9, snapshot['ami_name'], generic_cell)
            snapshots_worksheet.write(row, 10, snapshot['instance_id'], generic_cell)
            snapshots_worksheet.write(row, 11, snapshot['instance_name'], generic_cell)
            row += 1

if config["unattached_volumes"]["enabled"]:
    # Unattached Volumes
    print("\nLooking for unattached volumes")
    print("---")

    volumes_worksheet = workbook.add_worksheet("Unattached Volumes")
    row = 0
    # add headings
    volumes_worksheet.write(row, 0, "Volume ID", main_heading)
    volumes_worksheet.write(row, 1, "Create Time", main_heading)
    volumes_worksheet.write(row, 2, "Status", main_heading)
    volumes_worksheet.write(row, 3, "Size", main_heading)
    volumes_worksheet.write(row, 4, "Snapshot ID", main_heading)
    volumes_worksheet.write(row, 5, "Tags", main_heading)
    # set length of columns
    volumes_worksheet.set_column(0, 4, 30)
    volumes_worksheet.set_column(5, 5, 60)

    row += 1
    print("Fetching volumes...might take a while")
    for volume in get_available_volumes():
        print("Checking for volume", volume['id'])
        volumes_worksheet.write(row, 0, volume['id'], generic_cell)
        volumes_worksheet.write(row, 1, volume['create_time'], generic_cell)
        volumes_worksheet.write(row, 2, volume['status'], generic_cell)
        volumes_worksheet.write(row, 3, volume['size'], generic_cell)
        volumes_worksheet.write(row, 4, volume['snapshot_id'], generic_cell)
        volumes_worksheet.write(row, 5, volume['tags'], generic_cell)
        row += 1

if config["expensive_lambda_functions"]["enabled"]:
    # Top most expensive lambdas
    print("\nLooking for expensive lambda functions...")
    print("---")

    cost_percentage = config["expensive_lambda_functions"]["cost_percentage"]
    name_tag_key = config["expensive_lambda_functions"]["name_tag_key"]
    past_days = config["expensive_lambda_functions"]["past_days"]

    start_date = (datetime.today() - timedelta(days=past_days)).strftime("%Y-%m-%d")
    end_date = datetime.today().strftime("%Y-%m-%d")

    lambda_worksheet = workbook.add_worksheet("{}% Cost Lambdas".format(cost_percentage))
    row = 0
    # add headings
    lambda_worksheet.write(row, 0, "Function Name", main_heading)
    lambda_worksheet.write(row, 1, "Cost (in USD) for past {} days".format(past_days), main_heading)
    # set length of columns
    lambda_worksheet.set_column(0, 1, 60)
    row += 1

    client = boto3.client("ce")
    response = client.get_cost_and_usage(TimePeriod={'Start': start_date, 'End': end_date}, Granularity='DAILY', Metrics=['UnblendedCost'], GroupBy=[{'Type': 'TAG', 'Key': name_tag_key}], Filter={'Dimensions': {'Key': 'SERVICE', 'Values': ['AWS Lambda']}})
    raw_data = response.get('ResultsByTime', [])
    next_token = response.get('NextPageToken')
    while next_token is not None:
        response = client.get_cost_and_usage(TimePeriod={'Start': start_date, 'End': end_date}, Granularity='DAILY', Metrics=['UnblendedCost'], GroupBy=[{'Type': 'TAG', 'Key': name_tag_key}], Filter={'Dimensions': {'Key': 'SERVICE', 'Values': ['AWS Lambda']}}, NextPageToken=next_token)
        raw_data.extend(response['ResultsByTime'])
        next_token = response.get('NextPageToken')

    clean_data = defaultdict(float)
    total_cost = 0
    for interval in raw_data:
        for group_row in interval['Groups']:
            function_name = group_row['Keys'][0][len(name_tag_key)+1:]
            if function_name:
                interval_fn_cost = float(group_row['Metrics']['UnblendedCost']['Amount'])
                clean_data[function_name] += interval_fn_cost
                total_cost += interval_fn_cost

    sorted_functions = sorted(clean_data.items(), key=lambda x: x[1], reverse=True)

    target_amount = (cost_percentage / 100.0) * total_cost
    current_amount = 0
    for function_name, cost in sorted_functions:
        lambda_worksheet.write(row, 0, function_name, generic_cell)
        lambda_worksheet.write(row, 1, cost, generic_cell)
        row += 1
        current_amount += cost
        if current_amount > target_amount:
            break
    if total_cost > 0:
        lambda_worksheet.write(row, 0, "ALL FUNCTIONS", sub_heading)
        lambda_worksheet.write(row, 1, total_cost, sub_heading)
        row += 1

if config["expensive_kinesis_streams"]["enabled"]:
    # Top most expensive kinesis streams
    print("\nLooking for expensive kinesis streams...")
    print("---")

    cost_percentage = config["expensive_kinesis_streams"]["cost_percentage"]
    name_tag_key = config["expensive_kinesis_streams"]["name_tag_key"]
    past_days = config["expensive_kinesis_streams"]["past_days"]

    start_date = (datetime.today() - timedelta(days=past_days)).strftime("%Y-%m-%d")
    end_date = datetime.today().strftime("%Y-%m-%d")

    kinesis_worksheet = workbook.add_worksheet("{}% Cost Streams".format(cost_percentage))
    row = 0
    # add headings
    kinesis_worksheet.write(row, 0, "Kinesis Stream Name", main_heading)
    kinesis_worksheet.write(row, 1, "Number of Shards", main_heading)
    kinesis_worksheet.write(row, 2, "Cost (in USD) for past {} days".format(past_days), main_heading)
    # set length of columns
    kinesis_worksheet.set_column(0, 2, 60)
    row += 1

    client = boto3.client("ce")
    response = client.get_cost_and_usage(TimePeriod={'Start': start_date, 'End': end_date}, Granularity='DAILY', Metrics=['UnblendedCost'], GroupBy=[{'Type': 'TAG', 'Key': name_tag_key}], Filter={'Dimensions': {'Key': 'SERVICE', 'Values': ['Amazon Kinesis']}})
    raw_data = response.get('ResultsByTime', [])
    next_token = response.get('NextPageToken')
    while next_token is not None:
        response = client.get_cost_and_usage(TimePeriod={'Start': start_date, 'End': end_date}, Granularity='DAILY', Metrics=['UnblendedCost'], GroupBy=[{'Type': 'TAG', 'Key': name_tag_key}], Filter={'Dimensions': {'Key': 'SERVICE', 'Values': ['Amazon Kinesis']}}, NextPageToken=next_token)
        raw_data.extend(response['ResultsByTime'])
        next_token = response.get('NextPageToken')

    clean_data = defaultdict(float)
    total_cost = 0
    for interval in raw_data:
        for group_row in interval['Groups']:
            name = group_row['Keys'][0][len(name_tag_key)+1:]
            if name:
                interval_stream_cost = float(group_row['Metrics']['UnblendedCost']['Amount'])
                clean_data[name] += interval_stream_cost
                total_cost += interval_stream_cost

    sorted_streams = sorted(clean_data.items(), key=lambda x: x[1], reverse=True)

    target_amount = (cost_percentage / 100.0) * total_cost
    current_amount = 0
    for name, cost in sorted_streams:

        # also fetch number of shards now
        print("Fetching number of shards for", name)
        try:
            kinesis_client = boto3.client("kinesis")
            kinesis_response = kinesis_client.describe_stream(StreamName=name, Limit=100)
            no_of_shards = len(kinesis_response.get('StreamDescription', dict()).get('Shards', []))
            has_more_shards = kinesis_response.get('StreamDescription', dict()).get('HasMoreShards', False)
            if has_more_shards:
                no_of_shards = str(no_of_shards) + "+"  # number of shards more than 100!
        except Exception as e:
            print(e)
            continue

        kinesis_worksheet.write(row, 0, name, generic_cell)
        kinesis_worksheet.write(row, 1, no_of_shards, generic_cell)
        kinesis_worksheet.write(row, 2, cost, generic_cell)

        row += 1
        current_amount += cost
        if current_amount > target_amount:
            break
    if total_cost > 0:
        kinesis_worksheet.write(row, 0, "ALL STREAMS", sub_heading)
        kinesis_worksheet.write(row, 1, "", sub_heading)
        kinesis_worksheet.write(row, 2, total_cost, sub_heading)
        row += 1

if config["expensive_ddb"]["enabled"]:
    # Top most expensive dynamodb tables
    print("\nLooking for expensive dynamodb tables...")
    print("---")

    cost_percentage = config["expensive_ddb"]["cost_percentage"]
    name_tag_key = config["expensive_ddb"]["name_tag_key"]
    past_days = config["expensive_ddb"]["past_days"]

    start_date = (datetime.today() - timedelta(days=past_days)).strftime("%Y-%m-%d")
    end_date = datetime.today().strftime("%Y-%m-%d")

    worksheet = workbook.add_worksheet("{}% Cost DynamoDB Tables".format(cost_percentage))
    row = 0
    # add headings
    worksheet.write(row, 0, "DynamoDB Table Name", main_heading)
    worksheet.write(row, 1, "Billing Mode", main_heading)
    worksheet.write(row, 2, "Number of Items", main_heading)
    worksheet.write(row, 3, "Storage in GB", main_heading)
    worksheet.write(row, 4, "Cost (in USD) for past {} days".format(past_days), main_heading)
    # set length of columns
    worksheet.set_column(0, 5, 60)
    row += 1

    client = boto3.client("ce")
    response = client.get_cost_and_usage(TimePeriod={'Start': start_date, 'End': end_date}, Granularity='DAILY', Metrics=['UnblendedCost'], GroupBy=[{'Type': 'TAG', 'Key': name_tag_key}], Filter={'Dimensions': {'Key': 'SERVICE', 'Values': ['Amazon DynamoDB']}})
    raw_data = response.get('ResultsByTime', [])
    next_token = response.get('NextPageToken')
    while next_token is not None:
        response = client.get_cost_and_usage(TimePeriod={'Start': start_date, 'End': end_date}, Granularity='DAILY', Metrics=['UnblendedCost'], GroupBy=[{'Type': 'TAG', 'Key': name_tag_key}], Filter={'Dimensions': {'Key': 'SERVICE', 'Values': ['Amazon DynamoDB']}}, NextPageToken=next_token)
        raw_data.extend(response['ResultsByTime'])
        next_token = response.get('NextPageToken')
    clean_data = defaultdict(float)
    total_cost = 0
    for interval in raw_data:
        for group_row in interval['Groups']:
            name = group_row['Keys'][0][len(name_tag_key)+1:]
            if name:
                interval_cost = float(group_row['Metrics']['UnblendedCost']['Amount'])
                clean_data[name] += interval_cost
                total_cost += interval_cost

    sorted_tables = sorted(clean_data.items(), key=lambda x: x[1], reverse=True)

    target_amount = (cost_percentage / 100.0) * total_cost
    current_amount = 0
    for name, cost in sorted_tables:
        # fetch details for tables
        print("Fetching details for ", name)
        try:
            dynamodb = boto3.client("dynamodb")
            ddb_table = dynamodb.describe_table(TableName=name)['Table']
            number_of_items = ddb_table.get('ItemCount', 0)
            storage_in_gb = ddb_table.get('TableSizeBytes', 0) / 1024.0 / 1024.0 / 1024.0
            billing_mode = ddb_table.get('BillingModeSummary', dict()).get('BillingMode', "Not Available")
        except Exception as e:
            print(e)
            continue

        worksheet.write(row, 0, name, generic_cell)
        worksheet.write(row, 1, billing_mode, generic_cell)
        worksheet.write(row, 2, number_of_items, generic_cell)
        worksheet.write(row, 3, storage_in_gb, generic_cell)
        worksheet.write(row, 4, cost, generic_cell)

        row += 1
        current_amount += cost
        if current_amount > target_amount:
            break
    if total_cost > 0:
        worksheet.write(row, 0, "ALL TABLES", sub_heading)
        worksheet.write(row, 1, "", sub_heading)
        worksheet.write(row, 2, "", sub_heading)
        worksheet.write(row, 3, "", sub_heading)
        worksheet.write(row, 4, total_cost, sub_heading)
        row += 1

if config["on_demand_ddb"]["enabled"]:
    # On Demand DynamoDB Tables
    print("\nLooking for on demand dynamodb tables...")
    print("---")

    ddb_worksheet = workbook.add_worksheet("On-Demand DynamoDB Tables")
    row = 0
    # add headings
    ddb_worksheet.write(row, 0, "DynamoDB Table Name", main_heading)
    row += 1
    ddb_worksheet.set_column(0, 0, 40)

    client = boto3.client("dynamodb")
    for table in get_dynamodb_tables(client):
        print("Checking for table", table)
        billing_mode = client.describe_table(TableName=table)['Table'].get('BillingModeSummary', dict()).get('BillingMode', "")
        if billing_mode == "PAY_PER_REQUEST":
            # table is on-demand
            ddb_worksheet.write(row, 0, table)
            row += 1

# reusable values
log_group_gb = defaultdict(float)
past_days = 14
if config["storage_cloudwatch_log_groups"]["enabled"]:
    # Top N CloudWatch Log Groups by incoming bytes
    print("\nLooking for incoming bytes cloudwatch log groups...")
    print("---")

    top_n = config["storage_cloudwatch_log_groups"]["top_n"]
    past_days = config["storage_cloudwatch_log_groups"]["past_days"]

    cloudwatch_worksheet = workbook.add_worksheet("Top {} Log Groups".format(top_n))
    row = 0
    # add headings
    cloudwatch_worksheet.write(row, 0, "CloudWatch Log Group", main_heading)
    cloudwatch_worksheet.write(row, 1, "Incoming GBs in last {} days".format(past_days), main_heading)
    # set length of columns
    cloudwatch_worksheet.set_column(0, 1, 60)
    row += 1

    client = boto3.client("cloudwatch")

    start_time = int(datetime.timestamp(datetime.now() - timedelta(days=past_days)))
    end_time = int(datetime.timestamp(datetime.now()))
    period = end_time - start_time

    # fetch log groups active in last 14 days
    response = client.list_metrics(Namespace='AWS/Logs', MetricName="IncomingBytes")
    metrics = response.get('Metrics', [])
    next_token = response.get('NextToken')
    while next_token is not None:
        response = client.list_metrics(Namespace='AWS/Logs', MetricName="IncomingBytes")
        metrics.extend(response['Metrics'])
        next_token = response.get('NextToken')

    # form data queries in required format
    metric_data_queries = list()
    for metric in metrics:
        unique_id = generate_unique_string()
        metric_data_queries.append({
            "Id": unique_id,
            "MetricStat": {
                "Metric": metric,
                "Period": period,
                "Stat": "Sum",
                "Unit": "Bytes"
            }
        })

    # now fetch incoming bytes for them in batches
    results = []
    BATCH_SIZE = 100
    for index in range(0, len(metric_data_queries), BATCH_SIZE):
        response = client.get_metric_data(MetricDataQueries=metric_data_queries[index:index+BATCH_SIZE], StartTime=start_time, EndTime=end_time)
        batch_results = response.get('MetricDataResults', [])
        next_token = response.get('NextToken')
        while next_token is not None:
            response = client.get_metric_data(MetricDataQueries=metric_data_queries[index:index+BATCH_SIZE], StartTime=start_time, EndTime=end_time, NextToken=next_token)
            batch_results.extend(response['MetricDataResults'])
            next_token = response.get('NextToken')
        for result in batch_results:
            if result['StatusCode'] == "Complete" and result['Label'] != "IncomingBytes":
                try:
                    incoming_gb = result['Values'][0] / (1024 * 1024 * 1024)  # convert bytes to GB
                    results.append((result['Label'], incoming_gb))
                    log_group_gb[result['Label']] = incoming_gb
                except IndexError:
                    continue

    results.sort(key=lambda x: x[1], reverse=True)
    results = results[:top_n]

    for name, incoming_gb in results:
        cloudwatch_worksheet.write(row, 0, name, generic_cell)
        cloudwatch_worksheet.write(row, 1, incoming_gb, generic_cell)
        row += 1

if config['api_gateway_cloudwatch']["enabled"]:
    # Top N API Gateway REST API stages CloudWatch Log Groups

    if not config["storage_cloudwatch_log_groups"]["enabled"]:
        raise Exception("ERROR: Cannot add API Gateway sheet since storage_cloudwatch_log_groups was not enabled :(")

    print("\nLooking for API Gateway REST API stages CloudWatch Log Groups...")
    print("---")
    top_n = config["api_gateway_cloudwatch"]["top_n"]

    api_gateway_worksheet = workbook.add_worksheet("Top {} API GW Logs".format(top_n))
    row = 0
    # add headings
    api_gateway_worksheet.write(row, 0, "REST API", main_heading)
    api_gateway_worksheet.write(row, 1, "Stage", main_heading)
    api_gateway_worksheet.write(row, 2, "Execution Log Group", main_heading)
    api_gateway_worksheet.write(row, 3, "Incoming GBs in last {} days".format(past_days), main_heading)
    api_gateway_worksheet.write(row, 4, "Access Log Group", main_heading)
    api_gateway_worksheet.write(row, 5, "Incoming GBs in last {} days".format(past_days), main_heading)
    # set length of columns
    api_gateway_worksheet.set_column(0, 5, 30)
    row += 1

    client = boto3.client("apigateway")

    # get all apis first
    print("Fetching all REST APIs....")
    apis = []
    response = client.get_rest_apis()
    for item in response.get('items', []):
        apis.append({
            "api_id": item['id'],
            "api_name": item['name']
        })
    position = response.get('position')
    while position is not None:
        response = client.get_rest_apis(position=position)
        for item in response.get('items', []):
            apis.append({
                "api_id": item['id'],
                "api_name": item['name']
            })
        position = response.get('position')

    # check stages for REST APIs
    results = []
    for api in apis:
        print("Checking stages for", api['api_name'], "...")
        stages = client.get_stages(restApiId=api['api_id']).get('item', [])
        for stage in stages:
            access_log_group = stage.get('accessLogSettings', dict()).get('destinationArn', '')
            if access_log_group:
                access_log_group = access_log_group[access_log_group.rindex(":")+1:]
            access_usage = log_group_gb[access_log_group]
            log_group = "API-Gateway-Execution-Logs_{}/{}".format(api['api_id'], stage['stageName'])
            execution_usage = log_group_gb[log_group]
            if access_usage + execution_usage > 0:
                results.append([api['api_name'], stage['stageName'], log_group, log_group_gb[log_group], access_log_group, log_group_gb[access_log_group]])
    results.sort(key=lambda x: x[3]+x[5], reverse=True)
    results = results[:top_n]
    for result_row in results:
        for i in range(0, 6):
            api_gateway_worksheet.write(row, i, result_row[i], generic_cell)
        row += 1


if config["unused_elastic_ips"]["enabled"]:
    # Unused Elastic IPs
    print("\nLooking for unused elastic IPs...")
    print("---")

    elastic_ips_worksheet = workbook.add_worksheet("Unused Elastic IPs")
    row = 0
    # add headings
    elastic_ips_worksheet.write(row, 0, "Elastic Public IP", main_heading)
    elastic_ips_worksheet.write(row, 1, "Assigned to Instance", main_heading)
    elastic_ips_worksheet.write(row, 2, "Instance State", main_heading)
    elastic_ips_worksheet.write(row, 3, "Instance Name", main_heading)
    elastic_ips_worksheet.write(row, 4, "Instance ID", main_heading)
    # set length of columns
    elastic_ips_worksheet.set_column(0, 4, 30)
    row += 1

    client = boto3.client("ec2")
    addresses = client.describe_addresses()['Addresses']
    for address in addresses:
        instance_id = address.get('InstanceId', "")
        instance_name = ""
        instance_status = ""
        if instance_id:
            try:
                response = client.describe_instances(InstanceIds=[instance_id])
                instance_details = response['Reservations'][0]['Instances'][0]
            except Exception as e:
                print(e)
                continue
            for tag in instance_details['Tags']:
                if tag["Key"] == "Name":
                    instance_name = tag["Value"]
            instance_state = instance_details['State']['Name']
        if not instance_id or instance_state != "running":
            elastic_ips_worksheet.write(row, 0, address["PublicIp"], generic_cell)
            if not instance_id:
                elastic_ips_worksheet.write(row, 1, "NO", red_text_cell)
            else:
                elastic_ips_worksheet.write(row, 1, "YES", green_text_cell)
                elastic_ips_worksheet.write(row, 2, instance_state, generic_cell)
                elastic_ips_worksheet.write(row, 3, instance_name, generic_cell)
                elastic_ips_worksheet.write(row, 4, instance_id, generic_cell)
            row += 1

workbook.close()
