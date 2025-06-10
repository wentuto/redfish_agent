import argparse
import requests
import openpyxl
import json
import time
import re
from urllib.parse import urljoin
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

# ignore SSL warnings
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

def delay_function(endpoint):
    try:
        delay_seconds = int(endpoint)
        print(f"[*] Delay for {delay_seconds} seconds")
        for remaining in range(delay_seconds, 0, -1):
            print(f"    Remaining: {remaining} seconds", end='\r')
            time.sleep(1)
        print("    Delay complete. Moving to next command.\n")
    except ValueError:
        print(f"[!] Invalid delay time: {endpoint}")

def parse_change_username_endpoint(endpoint, username_to_id_map):
    pattern = r'\${([^.]+)\.id}'
    match = re.search(pattern, endpoint)
    if match:
        username_key = match.group(1)
        if username_key in username_to_id_map:
            endpoint = re.sub(pattern, username_to_id_map[username_key], endpoint)
            print(f"username_key: {username_key}, ID: {username_to_id_map[username_key]}")
            print(f"[*] Replaced dynamic endpoint with ID: {endpoint}")
        else:
            print(f"[!] Warning: No ID found for username {username_key}")
    return endpoint

def find_username_id(method, endpoint, status_code, response_json, username_to_id_map):
    # Store username to ID mapping for account creation
    check_then_store = False
    if (method.upper() == "POST" and 
        endpoint == "/redfish/v1/AccountService/Accounts" and 
        status_code == 201):
        check_then_store = True
        
    # Also store username to ID mapping from GET requests to specific account endpoints
    elif (method.upper() == "GET" and 
            endpoint.startswith("/redfish/v1/AccountService/Accounts/") and
            endpoint != "/redfish/v1/AccountService/Accounts/" and
            status_code == 200):
        check_then_store = True

    # Also store username to ID mapping from update operations (PATCH/PUT)
    elif ((method.upper() == "PATCH" or method.upper() == "PUT") and 
            endpoint.startswith("/redfish/v1/AccountService/Accounts/") and
            endpoint != "/redfish/v1/AccountService/Accounts/" and
            status_code in [200, 202, 204]):
        check_then_store = True
 
    if check_then_store:
        if "UserName" in response_json and "Id" in response_json:
            username_value = response_json["UserName"]
            id_value = response_json["Id"]
            username_to_id_map[username_value] = id_value
            print(f"[*] Stored mapping: {username_value} -> {id_value}")

def execute_redfish(username, password, root_url, excel_path='commands.xlsx', output_excel_path='output.xlsx'):
    try:
        wb = openpyxl.load_workbook(excel_path)
        sheet = wb.active

        # build Excel file for output
        output_wb = openpyxl.Workbook()
        output_sheet = output_wb.active
        output_sheet.append(["Method", "Endpoint", "Payload", "Status Code", "Response"])  # write title

        # Dictionary to store username to id mappings
        username_to_id_map = {}
        
        for col_num, column_title in enumerate(["Method", "Endpoint", "Payload", "Status Code", "Response"], 1):
            column_letter = get_column_letter(col_num)
            output_sheet.column_dimensions[column_letter].width = max(10, len(column_title) + 2)  # Minimum width of 10

        for row_num, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
            method, endpoint, payload = row

            if method.strip().lower() == "delay":
                delay_function(endpoint) 
                continue

            # Check if endpoint contains ${username.id} pattern and replace with the stored ID
            if "${" in endpoint and "id}" in endpoint:
                endpoint = parse_change_username_endpoint(endpoint, username_to_id_map) 

            url = urljoin(root_url, endpoint)
            headers = {"Content-Type": "application/json"}
            auth = (username, password)

            # 處理 payload
            data = None
            if payload:
                try:
                    data = json.loads(payload)
                except json.JSONDecodeError:
                    print(f"[!] invalid JSON payload: {payload}")
                    continue

            print(f"[*] {method} {url}")
            try:
                response = requests.request(
                    method=method,
                    url=url,
                    auth=auth,
                    headers=headers,
                    json=data,
                    verify=False,
                    timeout=10
                )
                status_code = response.status_code
                try:
                    response_json = response.json()
                    response_text = json.dumps(response_json, indent=4, ensure_ascii=False)

                    # Check if the response contains a username to ID mapping
                    find_username_id(method, endpoint, status_code, response_json, username_to_id_map)

                except json.JSONDecodeError:
                    response_text = response.text  
                print(f"Status Code: {status_code}")
                print(f"Response: {response_text}\n")

                output_sheet.append([method, endpoint, payload, status_code, response_text])

                # Adjust column widths based on content length
                for col_num, cell_value in enumerate([method, endpoint, payload, status_code, response_text], 1):
                    column_letter = get_column_letter(col_num)
                    current_width = output_sheet.column_dimensions[column_letter].width
                    output_sheet.column_dimensions[column_letter].width = max(current_width, len(str(cell_value)) + 2)
                    
                    # Enable text wrapping for the cell
                    cell = output_sheet.cell(row=row_num, column=col_num)
                    cell.alignment = Alignment(wrap_text=True)

                # Adjust row height dynamically
                num_lines = response_text.count('\n') + 1
                output_sheet.row_dimensions[row_num].height = 15 * num_lines

            except Exception as e:
                print(f"[!] request error: {e}")
                output_sheet.append([method, endpoint, payload, "Error", str(e)])

                # Adjust column widths even when there's an error
                for col_num, cell_value in enumerate([method, endpoint, payload, "Error", str(e)], 1):
                    column_letter = get_column_letter(col_num)
                    current_width = output_sheet.column_dimensions[column_letter].width
                    output_sheet.column_dimensions[column_letter].width = max(current_width, len(str(cell_value)) + 2)

                    # Enable text wrapping for the cell even when there's an error
                    cell = output_sheet.cell(row=row_num, column=col_num)
                    cell.alignment = Alignment(wrap_text=True)

                # Adjust row height dynamically even when there's an error
                output_sheet.row_dimensions[row_num].height = 15  # Set a default row height

        # 儲存輸出 Excel 檔案
        output_wb.save(output_excel_path)
        print(f"[*] Output result to：{output_excel_path}")

    except FileNotFoundError:
        print(f"[!] Excel file not found: {excel_path}")
    except Exception as e:
        print(f"[!] Execution error: {e}")

def main():
    parser = argparse.ArgumentParser(description="Execute Redfish commands from Excel.")
    parser.add_argument("-u", "--username", required=True, help="Username")
    parser.add_argument("-p", "--password", required=True, help="Password")
    parser.add_argument("-r", "--root", required=True, help="Root URL (e.g., https://127.0.0.1:5101)")
    parser.add_argument("-f", "--file", default="commands.xlsx", help="Excel file path")
    parser.add_argument("-o", "--output", default="output.xlsx", help="Output Excel file path") 

    args = parser.parse_args()

    execute_redfish(args.username, args.password, args.root, args.file, args.output)  

if __name__ == "__main__":
    main()

