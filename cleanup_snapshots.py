"""
Deletes all unreferenced Snapshots
"""

from ebs_helpers import get_snapshots, ec2
print("Fetching snapshots...")
for snapshot in get_snapshots():
    print("Checking", snapshot['id'])
    if not snapshot['volume_exists'] and not snapshot['ami_exists'] and not snapshot['instance_exists']:
        print("Deleting ", snapshot['id'])
        ec2.delete_snapshot(SnapshotId=snapshot['id'])