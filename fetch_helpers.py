"""
Helper functions to fetch details of different AWS Resources, can be used by other scripts
"""

def get_lambda_functions(client):
    # fetch all lambda functions info and store in "functions"
    fetched_functions = client.list_functions()
    next_marker = fetched_functions.get('NextMarker')
    functions = fetched_functions['Functions']
    while next_marker is not None:
        fetched_functions = client.list_functions(Marker=next_marker)
        functions.extend(fetched_functions['Functions'])
        next_marker = fetched_functions.get('NextMarker')
    return iter(functions)

def get_dynamodb_tables(client):
    # fetch all dynamodb tables and store in "tables"
    fetched_tables = client.list_tables()
    next_marker = fetched_tables.get('LastEvaluatedTableName')
    tables = fetched_tables['TableNames']
    while next_marker is not None:
        fetched_tables = client.list_tables(ExclusiveStartTableName=next_marker)
        next_marker = fetched_tables.get('LastEvaluatedTableName')
        tables.extend(fetched_tables['TableNames'])
    return iter(tables)

def get_ec2_reservations(client):
    # fetch all EC2 Instances info and store in "reservations"
    fetched_reservations = client.describe_instances()
    reservations = fetched_reservations['Reservations']
    next_token = fetched_reservations.get('NextToken')
    while next_token is not None:
        fetched_reservations = client.describe_instances(NextToken=next_token)
        reservations.extend(fetched_reservations['Reservations'])
        next_token = fetched_reservations.get('NextToken')
    return iter(reservations)

def get_kinesis_streams(client):
    # fetch all Kinesis streams info and store in "streams"
    fetched_streams = client.list_streams()
    streams = fetched_streams['StreamNames']
    has_more = fetched_streams.get('HasMoreStreams', False)
    while has_more == True:
        fetched_streams = client.list_streams(ExclusiveStartStreamName=streams[-1])
        streams.extend(fetched_streams['StreamNames'])
        has_more = fetched_streams.get('HasMoreStreams', False)
    return iter(streams)

def get_firehose_delivery_streams(client):
    # fetch all Firehose Delivery streams info and store in "streams"
    fetched_streams = client.list_delivery_streams()
    streams = fetched_streams['DeliveryStreamNames']
    has_more = fetched_streams.get('HasMoreDeliveryStreams', False)
    while has_more == True:
        fetched_streams = client.list_delivery_streams(ExclusiveStartDeliveryStreamName=streams[-1])
        streams.extend(fetched_streams['DeliveryStreamNames'])
        has_more = fetched_streams.get('HasMoreDeliveryStreams', False)
    return iter(streams)