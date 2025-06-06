## Description

This script, `exe_redfish.py`, executes Redfish commands from an Excel file and outputs the results to another Excel file.

## Installation

1.  **Install Python:** Ensure you have Python 3.6 or higher installed.

2.  **Clone the repository:**

    ```bash
    git clone git@github.com:wentuto/redfish_agent.git
    cd redfish_agent
    ```

3.  **Install dependencies:**

    ```bash
    pip install -r requirements.txt
    ```

    The `requirements.txt` file lists the necessary Python packages:

    *   `requests`: For making HTTP requests.
    *   `openpyxl`: For reading and writing Excel files.

## Usage

1.  **Prepare the Excel command file:**

    Create an Excel file (default name is `commands.xlsx`) with the following structure:

    *   The first row should contain headers (e.g., "Method", "Endpoint", "Payload").
    *   Each subsequent row represents a Redfish command.
        *   **Method:** HTTP method (e.g., GET, POST, PATCH, DELETE).
        *   **Endpoint:** Redfish endpoint (e.g., `/redfish/v1/Systems`).
        *   **Payload:** JSON payload for the request (can be empty).

2.  **Execute Command Example:**
     
    ```bash
    python3 redfish_agent.py -u user_name -p user_password -r https://127.0.0.1:5101
    ```

