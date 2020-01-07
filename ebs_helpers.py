"""
Helper functions to fetch details of different AWS EBS related resources
"""

import re
import boto3
from collections import defaultdict

ec2 = boto3.client("ec2")
volume_details = dict() # to memoize volume name, instance id and instance name against volume id
instance_details = dict() # to memoize instance name for an instance id
snapshot_to_ami = defaultdict(list) # stores AMI details against snapshot id

def update_amis():
    """
    Fetches all self owned AMIs and store the image info against snapshot id in "snapshot_to_image"
    """
    images = ec2.describe_images(Owners=['self'])['Images']
    for image in images:
        image_id = image['ImageId']
        image_name = image['Name']
        for mapping in image.get('BlockDeviceMappings', []):
            snapshot_id = mapping.get('Ebs', dict()).get('SnapshotId')
            if snapshot_id:
                if not image_name:
                    image_name = "No Name Set"
                snapshot_to_ami[snapshot_id].append({
                    "name": image_name,
                    "id": image_id
                })

def get_instance_name(instance_id):
    """
    Returns instance name for the given instance_id, returns blank string if instance doesn't exists,
    and "No Name Set" in case instance exists but has no name assigned

    This function also memoize the instance details
    """
    if instance_id in instance_details:
        return instance_details[instance_id]
    try:
        instance = ec2.describe_instances(InstanceIds=[instance_id])['Reservations'][0]['Instances'][0]
    except Exception as e:
        print(e)
        instance_details[instance_id] = ""
        return ""
    instance_details[instance_id] = "No Name Set"
    for tag in instance.get("Tags", []):
        if tag['Key'] == 'Name':
            instance_details[instance_id] = tag['Value']
            break
    return instance_details[instance_id]


def get_volume_details(volume_id):
    """
    Get Volume and Attached Instances Details for specified volume_id.
    In response, `instance_name` is blank string if instance doesn't exists
    `name` is blank string if volume doesn't exists

    This function also memoize the information
    """
    if volume_id in volume_details:
        return volume_details[volume_id]
    volume_detail = {
        "name": "",
        "instance_id": "",
        "instance_name": ""
    }
    # 
    try:
        # try fetching volume information
        volume = ec2.describe_volumes(VolumeIds=[volume_id])["Volumes"][0]
    except Exception as e:
        print(e)
        volume_details[volume_id] = volume_detail
        return volume_detail
    # store volume name if available
    volume_detail['name'] = "No Name Set"
    for tag in volume.get("Tags", []):
        if tag['Key'] == 'Name':
            volume_detail['name'] = tag['Value']
            break
    # get instance information if available
    instance_id = []
    instance_name = []
    for attachment in volume.get("Attachments", []):
        current_instance_id = attachment.get("InstanceId")
        if current_instance_id:
            instance_id.append(current_instance_id)
            instance_name.append(get_instance_name(current_instance_id))
    # store instance information
    volume_detail["instance_id"] = ", ".join(instance_id)
    volume_detail["instance_name"] = ", ".join(instance_name)
    volume_details[volume_id] = volume_detail
    return volume_detail

def get_snapshots():
    """
    Get all snapshots.
    """
    update_amis()
    snapshots = ec2.describe_snapshots(OwnerIds=['self'])['Snapshots']
    for snapshot in snapshots:
        snapshot_id = snapshot['SnapshotId']
        volume_id = snapshot['VolumeId']
        volume_details = get_volume_details(volume_id)
        ami_id = ""
        ami_exists = False
        ami_name = ""
        if snapshot_id in snapshot_to_ami:
            ami_exists = True
            ami_id = []
            ami_name = []
            for image in snapshot_to_ami[snapshot_id]:
                ami_name.append(image["name"])
                ami_id.append(image["id"])
            ami_id = ", ".join(ami_id)
            ami_name = ", ".join(ami_name)
        yield {
            'id': snapshot_id,
            'description': snapshot['Description'],
            'start_time': snapshot['StartTime'].strftime("%d-%m-%Y %H:%M:%S"),
            'size': snapshot['VolumeSize'],
            'volume_id': volume_id,
            'volume_exists': True if volume_details["name"] else False,
            'volume_name': volume_details["name"],
            'instance_id': volume_details["instance_id"],
            'instance_exists': True if volume_details["instance_name"] else False,
            'instance_name': volume_details["instance_name"],
            'ami_id': ami_id,
            'ami_exists': ami_exists,
            'ami_name': ami_name
        }


def get_available_volumes():
    """
    Get all volumes in 'available' state. (Volumes not attached to any instance)
    """
    volumes = ec2.describe_volumes(Filters=[{'Name': 'status', 'Values': ['available']}])['Volumes']
    for volume in volumes:
        tags = [ "{} = {}".format(tag['Key'], tag['Value']) for tag in volume.get('Tags', [])]
        yield {
            'id': volume['VolumeId'],
            'create_time': volume['CreateTime'].strftime("%d-%m-%Y %H:%M:%S"),
            'status': volume['State'],
            'size': volume['Size'],
            'snapshot_id': volume['SnapshotId'],
            'tags': ", ".join(tags),
        }