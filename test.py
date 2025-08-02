import win32com.client
import pandas as pd
import datetime
import subprocess
import time


# Parameters
COLLECTION_NAME = "Your Device Collection Name"  # <-- Change this
DAYS_THRESHOLD = 90
MAX_PING_ATTEMPTS = 3
EXCEL_OUTPUT_PATH = "stale_devices.xlsx"


# Connect to SCCM
def connect_sccm():
    sms_namespace = "root\\SMS\\site_YOUR_SITE_CODE"  # <-- Replace YOUR_SITE_CODE with your SCCM site code
    locator = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    sccm = locator.ConnectServer(".", sms_namespace)
    return sccm


# Query devices in a specific collection
def query_devices(sccm, collection_name):
    query = f"SELECT * FROM SMS_FullCollectionMembership WHERE CollectionID = (SELECT CollectionID FROM SMS_Collection WHERE Name='{collection_name}')"
    return sccm.ExecQuery(query)


# Get device details
def get_device_details(sccm, resource_id):
    query = f"SELECT Name, SerialNumber, LastHardwareScan, UserName FROM SMS_G_System_COMPUTER_SYSTEM WHERE ResourceID={resource_id}"
    results = sccm.ExecQuery(query)
    for result in results:
        return {
            "ComputerName": result.Name,
            "SerialNumber": result.SerialNumber if hasattr(result, 'SerialNumber') else '',
            "LastHardwareScan": get_last_hardware_scan(sccm, resource_id),
            "UserName": result.UserName if hasattr(result, 'UserName') else ''
        }
    return None


# Get Last Hardware Scan Date
def get_last_hardware_scan(sccm, resource_id):
    query = f"SELECT LastHardwareScan FROM SMS_G_System_WORKSTATION_STATUS WHERE ResourceID={resource_id}"