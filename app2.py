from zk import ZK, const
import json
from datetime import datetime, timedelta
from collections import defaultdict
import tkinter as tk
from tkinter import simpledialog, messagebox
import pandas as pd
from openpyxl import load_workbook
from openpyxl.workbook.protection import WorkbookProtection


def fetch_attendance_data(device_ip, port, start_date_str=None, end_date_str=None, timeout=30, password=0):
    """
    Connects to the ZKTeco device and fetches attendance data filtered by a date range.
    :param device_ip: IP address of the ZKTeco device
    :param port: Port number (default 4370)
    :param start_date_str: Start date in format YYYY-MM-DD (optional)
    :param end_date_str: End date in format YYYY-MM-DD (optional)
    :param timeout: Timeout for the connection
    :param password: Device password, if set
    :return: Processed attendance data
    """
    conn = None
    try:
        zk = ZK(device_ip, port=port, timeout=timeout, password=password, force_udp=False, ommit_ping=False)
        conn = zk.connect()

        # Convert start and end dates to datetime objects
        if start_date_str:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        else:
            start_date = datetime.min  # If no start date is provided, consider the earliest possible date

        if end_date_str:
            end_date = datetime.strptime(end_date_str, "%Y-%m-%d") + timedelta(days=1)  # Include the end date
        else:
            end_date = datetime.now()  # If no end date is provided, consider the current date and time

        # Fetch attendance logs
        logs = conn.get_attendance()
        attendance_data = defaultdict(list)

        # Group logs by user_id and date
        for log in logs:
            log_timestamp = log.timestamp if isinstance(log.timestamp, datetime) else datetime.strptime(log.timestamp, "%Y-%m-%d %H:%M:%S")
            if start_date <= log_timestamp < end_date:
                user_date_key = (log.user_id, log_timestamp.date())
                attendance_data[user_date_key].append({
                    "user_id": log.user_id,
                    "timestamp": log_timestamp,
                    "status": log.status,
                    "punch": log.punch
                })

        # Process the grouped data to add check-in/check-out labels
        result_data = []
        for (user_id, date), logs in attendance_data.items():
            logs.sort(key=lambda x: x["timestamp"])
            if logs:
                first_punch = logs[0]
                last_punch = logs[-1]
                result_data.append({
                    "Employee": first_punch["user_id"],
                    "Time": first_punch["timestamp"].strftime("%Y-%m-%d %H:%M:%S"),
                    "Log Type": "IN"
                })
                result_data.append({
                    "Employee": last_punch["user_id"],
                    "Time": last_punch["timestamp"].strftime("%Y-%m-%d %H:%M:%S"),
                    "Log Type": "OUT"
                })

        # Return the processed data
        return result_data
    except Exception as e:
        print(f"Error: {e}")
        messagebox.showerror("Error", f"An error occurred: {e}")
        return []
    finally:
        if conn:
            conn.enable_device()
            conn.disconnect()

def save_to_protected_excel(data, filename="attendance_data.xlsx", sheet_password="secure"):
    """
    Saves the attendance data to a protected Excel file.
    :param data: List of dictionaries containing attendance data
    :param filename: Filename to save the data (default is 'attendance_data.xlsx')
    :param sheet_password: Password to protect the sheet (default is 'secure')
    """
    # Save data to Excel file
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)

    # Load the workbook to protect it
    wb = load_workbook(filename)
    sheet = wb.active

    # Protect the sheet
    sheet.protection.sheet = True
    sheet.protection.password = sheet_password

    # Protect the workbook structure (optional)
    wb.security = WorkbookProtection(workbookPassword=sheet_password, lockStructure=False)

    # Save the protected workbook
    wb.save(filename)
    print(f"Data saved to {filename} and protected with password.")
    messagebox.showinfo("Success", f"Data saved to {filename} and protected with password.")

# GUI for start and end date input
if __name__ == "__main__":
    # Initialize tkinter
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window

    # Prompt for device IP and date range
    device_ip = "192.168.1.201"
    port = 4370  # Default port for ZKTeco
    start_date = simpledialog.askstring("Input", "Enter the start date (YYYY-MM-DD):")
    end_date = simpledialog.askstring("Input", "Enter the end date (YYYY-MM-DD):")

    if device_ip and start_date and end_date:
        # Fetch attendance data
        attendance_data = fetch_attendance_data(device_ip, port, start_date_str=start_date, end_date_str=end_date)

        # Print the data to the terminal
        print(json.dumps(attendance_data, indent=4))

        # Save the data to a protected Excel file
        save_to_protected_excel(attendance_data, sheet_password="secure123")
    else:
        messagebox.showwarning("Warning", "All inputs are required!")
