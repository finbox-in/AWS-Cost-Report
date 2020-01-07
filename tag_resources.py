"""
Tags resources with specified tag values in form of a CSV File 

Usage
-------
Expects a file to_tag.csv with columns Resource,Name,Tag:Name,Tag:<tagname1>,Tag:<tagname1>,...
Make sure the column names are present on the first row in the sequence as stated above
Here Resource can be: DynamoDB Table, Firehose Delivery Stream, Kinesis Stream, S3 Bucket, Lambda Function
"""

import boto3
import csv

def tag_resources(resources, batch_size=20):
    """
    resources -- list of resources dictionary having keys "arn" (string) and "tags" (dict)
    batch_size -- batch size for calling tag resources API (default is 20)
    """
    client = boto3.client('resourcegroupstaggingapi')

    for index in range(0, len(resources), batch_size):
        batch = resources[index:index+batch_size]
        for resource in batch:
            response = client.tag_resources(
                ResourceARNList=[
                    resource['arn'],
                ],
                Tags=resource['tags']
            )
            print(resource['arn'], response['ResponseMetadata']['HTTPStatusCode'])

def get_arn(resource_type, name):
    """
    Returns ARN for the given resource type and name
    """
    if resource_type == "DynamoDB Table":
        client = boto3.client("dynamodb")
        return client.describe_table(TableName=name)['Table']['TableArn']
    elif resource_type == "Firehose Delivery Stream":
        client = boto3.client("firehose")
        return client.describe_delivery_stream(DeliveryStreamName=name)['DeliveryStreamDescription']['DeliveryStreamARN']
    elif resource_type == "Kinesis Stream":
        client = boto3.client("kinesis")
        return client.describe_stream(StreamName=name)['StreamDescription']['StreamARN']
    elif resource_type == "S3 Bucket":
        return "arn:aws:s3:::{}".format(name)
    elif resource_type == "Lambda Function":
        client = boto3.client("lambda")
        return client.get_function(FunctionName=name)['Configuration']['FunctionArn']
    return None

with open("to_tag.csv") as csv_file:
    csv_reader = csv.reader(csv_file)
    line_count = 0
    resources = []
    tag_keys = []
    print("Reading from csv...might take a while")
    for row in csv_reader:
        if line_count == 0:
            for index in range(0, len(row)-2):
                # length of `Tag:` is 4
                tag_keys.append(row[index+2][4:])
            print("Tags to be added :", ",".join(tag_keys))
        else:
            arn = get_arn(row[0], row[1])
            if arn:
                tags = dict()
                for index, value in enumerate(tag_keys):
                    tags[value] = row[index+2]
                resource = {"arn": arn, "tags": tags}
                resources.append(resource)
        line_count += 1
    print("Tagging now...")
    tag_resources(resources)
