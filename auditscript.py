import subprocess # Enables running external system commands and programs from within Python.
                  # Useful for executing shell commands (like 'net accounts', 'gpresult', etc.) and capturing their output.

import re # Provides Regular Expression (regex) support for advanced string searching and pattern matching.
          # Often used to extract specific information from command outputs or text.
from openpyxl import Workbook

def checks():
    data_list=[]
    # get host name
    try:
        tag_id = subprocess.check_output("hostname", shell=True, text=True).strip()
        print(tag_id)
        data_list.append(tag_id)
    except:
         print("Couldn't find computer name")
         data_list.append("Error")

    # get MAC Address(s)
    try:
        mac_add = subprocess.check_output('wmic nic where "PNPDeviceID like \'PCI%\'" get MACAddress',shell=True, text=True).strip().split('\n')[2:]
        print(mac_add)
        data_list.append(mac_add)
    except:
        print("Couldn't find MAC Address")
        data_list.append("Error") 
    
    # get HDD Serial Number from cmd
    try:
        hdd_serial = subprocess.check_output("wmic diskdrive get serialnumber", shell=True, text=True).strip().split('\n')[2:]
        print(hdd_serial)
        data_list.append(hdd_serial)
    except:
        print("Error! Couldn't find HDD Serial Number from cmd")
        data_list.append("Error")


    # get HDD Serial Number from power shell
    try:
        hdd_serial_ps = ["powershell","-Command","Get-CimInstance Win32_PhysicalMedia | Select-Object -ExpandProperty SerialNumber"]
        hdd_serial_ps_out = subprocess.run(hdd_serial_ps,capture_output=True,text=True,shell=False).stdout.strip()
        print(hdd_serial_ps_out)
        data_list.append(hdd_serial_ps_out)
    except:
        print("Error! Couldn't find HDD Serial Number from powershell")
        data_list.append("Error")
    
    # get USB Block Status from regedit
    try: 
        usb_block = subprocess.check_output(r'reg query HKLM\SYSTEM\CurrentControlSet\Services\USBSTOR /v Start', shell=True, text=True)
        if "0x4" in usb_block.lower():
            print("USB storage is BLOCKED")
            data_list.append(1)
        elif "0x3" in usb_block.lower():
            print("Alert! USB storage is ENABLED")
            data_list.append(0)
        else:
            print("Error! Unknown USB policy")
            data_list.append("Error")
    except:
        print("Error! Couldn't find USB Block Status from regedit")
        data_list.append("Error")

    # get Removable Storage Devices Access Status set from gpedit
    try:
        removable_disk_block = subprocess.check_output(r'reg query HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices\{53f5630d-b6bf-11d0-94f2-00a0c91efb8b}', shell=True, text=True).strip()  
        Deny_Read = re.search(r'Deny_Read\s+REG_DWORD\s+(\w+)', removable_disk_block)
        Deny_Write = re.search(r'Deny_Write\s+REG_DWORD\s+(\w+)', removable_disk_block)
        Deny_Execute = re.search(r'Deny_Execute\s+REG_DWORD\s+(\w+)', removable_disk_block)
        if "0x0" in [Deny_Read.group(1), Deny_Write.group(1), Deny_Execute.group(1)]:
                print("Alert! Removable Storage Devices Access Deny is DISABLED")
                data_list.append(0)
        elif "0x0" not in [Deny_Read.group(1), Deny_Write.group(1), Deny_Execute.group(1)]:
                print("Removable Storage Devices Access Deny is Enabled")
                data_list.append(1)
        else:
            print("Error! Unknown Removable Storage Devices policy")
            data_list.append("Error")
    except:
        print("Error! Couldn't find Removable Storage Devices Access Status")
        data_list.append("Error")

    service_names = ["Bluettooth Audio Gateway Service", 
                    "Bluetooth Support Service", 
                    "Fax",
                    "Remote Access Auto Connect Manager", 
                    "Remote Access Auto Manager",
                    "Remote Desktop Configuration", 
                    "Remote Desktop Services",
                    "Remote Desktop Services UserMode Port Redirector",
                    "Wi-Fi Direct Services Connection Manager Service",
                    "Windows Mobile Hotspot Service"]
    services = ["BTAGService", "BthServ", "Fax", "RasAuto", "RasMan", "SessionEnv", "TermService", "UmRdpService", "WdiServiceHost", "WlanSvc"]
    for service, service_name in zip(services, service_names):
        try:     
            service_check = subprocess.run(["sc", "qc", service], capture_output=True, text=True).stdout
            service_status = re.search(r'START_TYPE\s+:\s(\d)\s+\w+', service_check)
            if service_status.group(1) == "4":
                print(f"{service_name} is disabled")
                data_list.append(1) 
            elif service_status.group(1) != "4": 
                print(f"Alert! {service_name} is enabled")
                data_list.append(0)
            else:  
                print(f"Error! {service_name} not found")
                data_list.append("Error")
        except:
            print(f"Error! Couldn't find {service_name}")
            data_list.append("Error")

    try:
        usb_traces = subprocess.check_output(r'reg query HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\USBSTOR', shell=True, text=True).strip().split('\n')
        if usb_traces:
            traces_list = []
            for i in usb_traces:
                usb_trace = re.search(r'HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Enum\\USBSTOR\\(.*)', i)
                traces_list.append(usb_trace.group(1))
            data_list.append(traces_list)
            print(f"USB Traces: {traces_list}")
        else:
            print("No USB Traces Found")
            data_list.append("No USB Traces Found")
    except:
        print("Error! Couldn't find USB Traces")
        data_list.append("Error")

    try:
        usb_traces = subprocess.check_output(r'reg query HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\SWD\WPDBUSENUM', shell=True, text=True).strip().split('\n')
        if usb_traces:
            traces_list = []
            for i in usb_traces:
                usb_trace = re.search(r'HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Enum\\SWD\\WPDBUSENUM\\(.*)', i)
                traces_list.append(usb_trace.group(1))
            data_list.append(traces_list)
            print(f"USB Traces: {traces_list}")
        else:
            print("No USB Traces Found")
            data_list.append("No USB Traces Found")
    except:
        print("Error! Couldn't find USB Traces")
        data_list.append("Error")

    wb = Workbook()
    ws = wb.active

    # Write list horizontally in row 1
    for col, value in enumerate(data_list, start=1):
        ws.cell(row=1, column=col, value=str(value))
    
    # Protect the sheet with password
    ws.protection.sheet = True
    ws.protection.set_password("winter!$coming")
    
    # Save file
    wb.save(f"{tag_id}.xlsx")

    print("\nProgram finished successfully!")

    # keep console open
    input("Press ENTER to exit...")


def main():
    checks()

if __name__ == "__main__":
    main()