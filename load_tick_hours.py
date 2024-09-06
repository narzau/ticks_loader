import argparse
import pandas as pd
import requests

from bs4 import BeautifulSoup
from datetime import datetime
from urllib.parse import urlencode


LOGIN_URL = "https://secure.tickspot.com/login"
CREATE_TIMECARD_ENTRY_URL = "https://intermedia1.tickspot.com/timecard/entries"

project_name_to_id_map = {
    "MTK": "17471389"  # this id is Tick's task_id. which represents MTK
}


class LoginFailedException(Exception):
    pass


def get_token_from_login_response(login_response: requests.Response) -> str:
    soup = BeautifulSoup(login_response.text, "html.parser")
    csrf_meta = soup.find("meta", attrs={"name": "csrf-token"})

    if not csrf_meta:
        print("CSRF token meta tag not found")
        raise LoginFailedException("CSRF Token Not Found")

    return csrf_meta.get("content")


def login(username, password):
    payload = {
        "user_login": username,
        "user_password": password,
        "commit": "Sign In",
        "remember[password]": "1",
    }
    response = requests.post(LOGIN_URL, data=urlencode(payload))
    if response.status_code != 200:
        raise LoginFailedException("Login Failed")

    return response


def create_timecard_entry(dates, headers, login_response, project_name):
    for date in dates:
        formatted_date = datetime.strptime(date, "%d/%m/%Y").strftime("%Y-%m-%d")

        payload = {
            "entry[id]": "",
            "timer[id]": "",
            "entry[date]": formatted_date,
            "task[id]": project_name_to_id_map[project_name],
            "entry[hours]": "8",
            "entry[notes]": "",
            "commit": "Enter Time",
        }

        try:
            response = requests.post(
                CREATE_TIMECARD_ENTRY_URL,
                headers=headers,
                data=urlencode(payload),
                cookies=login_response.cookies,
            )
            if response.status_code == 200:
                print(f"Successful request for the date {formatted_date}")
            else:
                print(
                    f"Request error for the date {formatted_date}. Status code: {response.status_code}"
                )
        except requests.exceptions.RequestException as e:
            print(f"Error making the request for the date {formatted_date}: {e}")


def get_dates_from_excel(file_path, sheet_name, start_date, end_date):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        if "date" not in df.columns:
            print(f"Error: 'date' column not found in the sheet '{sheet_name}'")
            return []

        df["date"] = pd.to_datetime(df["date"], format="%d/%m/%Y", errors="coerce")

        df = df.dropna(subset=["date"])

        start_date_parsed = pd.to_datetime(start_date, format="%d/%m/%Y")
        end_date_parsed = pd.to_datetime(end_date, format="%d/%m/%Y")

        filtered_dates = df[
            (df["date"] >= start_date_parsed) & (df["date"] <= end_date_parsed)
        ]

        valid_dates = filtered_dates["date"].dt.strftime("%d/%m/%Y").tolist()
        return valid_dates
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return []


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Submit timecard entries.")
    parser.add_argument(
        "--project", type=str, default="MTK", help="Name of the project (default: MTK)"
    )
    parser.add_argument(
        "--hours", type=int, default=8, help="Number of hours to load (default: 8)"
    )
    parser.add_argument(
        "--email", type=str, required=True, help="Email to login into Tick"
    )
    parser.add_argument(
        "--password", type=str, required=True, help="Password to login into Tick"
    )
    parser.add_argument(
        "--file",
        type=str,
        required=True,
        help="Path to the Excel file containing dates",
    )
    parser.add_argument(
        "--sheet", type=str, required=True, help="Sheet name representing the person"
    )
    parser.add_argument(
        "--start_date",
        type=str,
        required=True,
        help="Start date for filtering in dd/mm/yyyy format",
    )
    parser.add_argument(
        "--end_date",
        type=str,
        required=True,
        help="End date for filtering in dd/mm/yyyy format",
    )
    args = parser.parse_args()

    if args.project not in project_name_to_id_map:
        print(
            f"Error: Unknown project '{args.project}'. Please provide a valid project. \nValid projects: ({','.join([project_name for project_name in project_name_to_id_map.keys()])})"
        )
        exit(1)

    dates = get_dates_from_excel(args.file, args.sheet, args.start_date, args.end_date)

    if not dates:
        print("No valid dates found in the specified date range. Exiting.")
        exit(1)

    print(
        f"This script is going to load {args.hours} hours per day for the {args.project} project."
    )
    print(f"Filtering dates from {args.start_date} to {args.end_date}.")
    print("The following dates will be processed:")
    for date in dates:
        print(f" - {date}")

    confirmation = input("Do you want to proceed? (yes/no): ").strip().lower()

    if confirmation != "yes":
        print("Operation cancelled.")
        exit(0)

    try:
        login_response = login(args.email, args.password)
    except Exception as e:
        print(f"Error during login: {e}")
        exit(1)

    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "x-csrf-token": get_token_from_login_response(login_response),
    }

    # create_timecard_entry(dates, headers, login_response, args.project)
