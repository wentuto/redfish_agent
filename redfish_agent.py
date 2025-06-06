import argparse
import requests
import openpyxl
import json
from urllib.parse import urljoin
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

# ignore SSL warnings
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

def execute_redfish(username, password, root_url, excel_path='commands.xlsx', output_excel_path='output.xlsx'):
    try:
        wb = openpyxl.load_workbook(excel_path)
        sheet = wb.active

        # build Excel file for output
        output_wb = openpyxl.Workbook()
        output_sheet = output_wb.active
        output_sheet.append(["Method", "Endpoint", "Payload", "Status Code", "Response"])  # write title

        # Dynamically adjust column widths based on header length
        for col_num, column_title in enumerate(["Method", "Endpoint", "Payload", "Status Code", "Response"], 1):
            column_letter = get_column_letter(col_num)
            output_sheet.column_dimensions[column_letter].width = max(10, len(column_title) + 2)  # Minimum width of 10

        for row_num, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
            method, endpoint, payload = row

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

