"""
Generates report.xlsx file, based on the options specified on config.json file
"""

import boto3
import xlsxwriter
import json
from datetime import datetime, timedelta
from ebs_helpers import get_snapshots, get_available_volumes
from fetch_helpers import get_lambda_functions, get_dynamodb_tables, get_ec2_reservations
from fetch_helpers import get_kinesis_streams, get_firehose_delivery_streams
from collections import defaultdict

workbook = xlsxwriter.Workbook('report.xlsx')
blue_heading = workbook.add_format({ 'font_color': '#ffffff', 'bg_color': '#0080ff', 'valign': 'vcenter', 'border': 1, 'font_size': 13 }) 
gray_heading = workbook.add_format({ 'font_color': '#ffffff', 'bg_color': '#969696', 'valign': 'vcenter', 'border': 1, 'font_size': 13 })
generic_cell = workbook.add_format({ 'valign': 'vcenter', 'border': 1, 'font_size': 13 })
green_text_cell = workbook.add_format({ 'valign': 'vcenter', 'border': 1, 'font_size': 13, 'font_color': '#037d50' })
red_text_cell = workbook.add_format({ 'valign': 'vcenter', 'border': 1, 'font_size': 13, 'font_color': '#cc0000' })

print("Loading config.json file....")
config = dict()
with open("config.json") as json_file:
    config = json.load(json_file)
print("configuration loaded!")

def generate_unique_string():
    UNIQUE_ID_SIZE = 10
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
                if value == True:
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


##############################################
##############################################
####     Main Code starts here            ####
##############################################
##############################################

if config["expensive_services"]["enabled"] == True:
    cost_percentage = config["expensive_services"]["cost_percentage"]
    past_days = config["expensive_services"]["past_days"]
    ##############################################
    ##       Expensive Services                 ##
    ##############################################
    print("\nLooking for expensive services")
    print("---")
    start_date = (datetime.today() - timedelta(days=past_days)).strftime("%Y-%m-%d")
    end_date = datetime.today().strftime("%Y-%m-%d")

    services_worksheet = workbook.add_worksheet("{}% Cost Services".format(cost_percentage))
    row = 0
    # add headings
    services_worksheet.write(row, 0, "Service", blue_heading)
    services_worksheet.write(row, 1, "Cost (in USD)", blue_heading)
    # set length of columns
    services_worksheet.set_column(0, 1, 60)
    row += 1


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
        services_worksheet.write(row, 0, "ALL SERVICES", gray_heading)
        services_worksheet.write(row, 1, total_cost, gray_heading)
        row += 1


if config["untagged_resources"]["enabled"] == True:
    tags_to_look = config["untagged_resources"]["tags"]
    ##############################################
    ##       Untagged Resources                 ##
    ##############################################
    print("\nLooking for untagged resources")
    print("---")

    untagged_worksheet = workbook.add_worksheet("Untagged Resources")
    row = 0
    # add headings
    untagged_worksheet.write(row, 0, "Resource", blue_heading)
    untagged_worksheet.write(row, 1, "Name", blue_heading)
    col = 2
    for key in tags_to_look:
        untagged_worksheet.write(row, col, key, blue_heading)
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
                instance_identifier += (" (" + tags['Name']+ ")")
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
        while has_more_tags == True:
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
        while has_more_tags == True:
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

if config["unreferenced_snapshots"]["enabled"] == True:
    ##############################################
    ##       Unreferenced Snapshots             ##
    ##############################################
    print("\nLooking for unreferenced snapshots")
    print("---")

    snapshots_worksheet = workbook.add_worksheet("Unreferenced Snapshots")
    row = 0
    # add headings
    snapshots_worksheet.write(row, 0, "Snapshot ID", blue_heading)
    snapshots_worksheet.write(row, 1, "Size", blue_heading)
    snapshots_worksheet.write(row, 2, "Start Time", blue_heading)
    snapshots_worksheet.write(row, 3, "Volume", blue_heading)
    snapshots_worksheet.write(row, 4, "AMI", blue_heading)
    snapshots_worksheet.write(row, 5, "Instance", blue_heading)
    snapshots_worksheet.write(row, 6, "Volume ID", blue_heading)
    snapshots_worksheet.write(row, 7, "Volume Name", blue_heading)
    snapshots_worksheet.write(row, 8, "AMI ID", blue_heading)
    snapshots_worksheet.write(row, 9, "AMI Name", blue_heading)
    snapshots_worksheet.write(row, 10, "Instance ID", blue_heading)
    snapshots_worksheet.write(row, 11, "Instance Name", blue_heading)
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

if config["unattached_volumes"]["enabled"] == True:
    ##############################################
    ##       Unattached Volumes                 ##
    ##############################################
    print("\nLooking for unattached volumes")
    print("---")

    volumes_worksheet = workbook.add_worksheet("Unattached Volumes")
    row = 0
    # add headings
    volumes_worksheet.write(row, 0, "Volume ID", blue_heading)
    volumes_worksheet.write(row, 1, "Create Time", blue_heading)
    volumes_worksheet.write(row, 2, "Status", blue_heading)
    volumes_worksheet.write(row, 3, "Size", blue_heading)
    volumes_worksheet.write(row, 4, "Snapshot ID", blue_heading)
    volumes_worksheet.write(row, 5, "Tags", blue_heading)
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

if config["expensive_lambda_functions"]["enabled"] == True:
    ##############################################
    ##       Top most expensive lambdas         ##
    ##############################################
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
    lambda_worksheet.write(row, 0, "Function Name", blue_heading)
    lambda_worksheet.write(row, 1, "Cost (in USD)", blue_heading)
    # set length of columns
    lambda_worksheet.set_column(0, 1, 60)
    row += 1

    client = boto3.client("ce")
    response = client.get_cost_and_usage(TimePeriod={'Start': start_date, 'End': end_date}, Granularity='DAILY', Metrics=['UnblendedCost'], GroupBy=[{'Type': 'TAG', 'Key': name_tag_key}], Filter={'Dimensions': { 'Key': 'SERVICE', 'Values': ['AWS Lambda']}})
    raw_data = response.get('ResultsByTime', [])
    next_token = response.get('NextPageToken')
    while next_token is not None:
        response = client.get_cost_and_usage(TimePeriod={'Start': start_date, 'End': end_date}, Granularity='DAILY', Metrics=['UnblendedCost'], GroupBy=[{'Type': 'TAG', 'Key': name_tag_key}], Filter={'Dimensions': { 'Key': 'SERVICE', 'Values': ['AWS Lambda']}}, NextPageToken=next_token)
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
        lambda_worksheet.write(row, 0, "ALL FUNCTIONS", gray_heading)
        lambda_worksheet.write(row, 1, total_cost, gray_heading)
        row += 1

if config["expensive_kinesis_streams"]["enabled"] == True:
    ##############################################
    ##    Top most expensive kinesis streams    ##
    ##############################################
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
    kinesis_worksheet.write(row, 0, "Kinesis Stream Name", blue_heading)
    kinesis_worksheet.write(row, 1, "Cost (in USD)", blue_heading)
    # set length of columns
    kinesis_worksheet.set_column(0, 1, 60)
    row += 1

    client = boto3.client("ce")
    response = client.get_cost_and_usage(TimePeriod={'Start': start_date, 'End': end_date}, Granularity='DAILY', Metrics=['UnblendedCost'], GroupBy=[{'Type': 'TAG', 'Key': name_tag_key}], Filter={'Dimensions': { 'Key': 'SERVICE', 'Values': ['Amazon Kinesis']}})
    raw_data = response.get('ResultsByTime', [])
    next_token = response.get('NextPageToken')
    while next_token is not None:
        response = client.get_cost_and_usage(TimePeriod={'Start': start_date, 'End': end_date}, Granularity='DAILY', Metrics=['UnblendedCost'], GroupBy=[{'Type': 'TAG', 'Key': name_tag_key}], Filter={'Dimensions': { 'Key': 'SERVICE', 'Values': ['Amazon Kinesis']}}, NextPageToken=next_token)
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
        kinesis_worksheet.write(row, 0, name, generic_cell)
        kinesis_worksheet.write(row, 1, cost, generic_cell)
        row += 1
        current_amount += cost
        if current_amount > target_amount:
            break
    if total_cost > 0:
        kinesis_worksheet.write(row, 0, "ALL STREAMS", gray_heading)
        kinesis_worksheet.write(row, 1, total_cost, gray_heading)
        row += 1

if config["storage_cloudwatch_log_groups"]["enabled"] == True:
    ########################################################
    ##    Top N CloudWatch Log Groups by incoming bytes   ##
    ########################################################
    print("\nLooking for incoming bytes cloudwatch log groups...")
    print("---")

    top_n = config["storage_cloudwatch_log_groups"]["top_n"]
    past_days = config["storage_cloudwatch_log_groups"]["past_days"]

    cloudwatch_worksheet = workbook.add_worksheet("Top {} Log Groups".format(top_n))
    row = 0
    # add headings
    cloudwatch_worksheet.write(row, 0, "CloudWatch Log Group", blue_heading)
    cloudwatch_worksheet.write(row, 1, "Incoming GBs in last {} days".format(past_days), blue_heading)
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
                    incoming_gb = result['Values'][0] / (1024 * 1024 * 1024) # convert bytes to GB
                    results.append((result['Label'], incoming_gb))
                except IndexError:
                    continue

    results.sort(key = lambda x: x[1], reverse=True)
    
    for name, incoming_gb in results[:top_n]:
        cloudwatch_worksheet.write(row, 0, name, generic_cell)
        cloudwatch_worksheet.write(row, 1, incoming_gb, generic_cell)
        row += 1

if config["unused_elastic_ips"]["enabled"] == True:
    ########################################################
    ##           Unused Elastic IPs                       ##
    ########################################################
    print("\nLooking for unused elastic IPs...")
    print("---")

    elastic_ips_worksheet = workbook.add_worksheet("Unused Elastic IPs")
    row = 0
    # add headings
    elastic_ips_worksheet.write(row, 0, "Elastic Public IP", blue_heading)
    elastic_ips_worksheet.write(row, 1, "Assigned to Instance", blue_heading)
    elastic_ips_worksheet.write(row, 2, "Instance State", blue_heading)
    elastic_ips_worksheet.write(row, 3, "Instance Name", blue_heading)
    elastic_ips_worksheet.write(row, 4, "Instance ID", blue_heading)
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