from pathlib import Path
import pandas as pd
import logging
from datetime import datetime

# -----------------------------
# File paths
# -----------------------------
BASE_DIR = Path(__file__).resolve().parent

INPUT_FILE = BASE_DIR / "employees.csv"
CLEANED_FILE = BASE_DIR / "cleaned_employees.xlsx"
DEPARTMENT_FILE = BASE_DIR / "department_summary.xlsx"
ERROR_FILE = BASE_DIR / "validation_errors.csv"
SUMMARY_FILE = BASE_DIR / "summary_report.txt"
LOG_FILE = BASE_DIR / "processing_log.txt"

# -----------------------------
# Logging
# -----------------------------
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)


def read_employee_data():
    try:
        df = pd.read_csv(INPUT_FILE)
        logging.info("Employee data file read successfully.")
        return df

    except FileNotFoundError:
        logging.error("employees.csv file was not found.")
        raise

    except Exception as e:
        logging.error(f"Error while reading input file: {e}")
        raise


def validate_columns(df):
    required_columns = [
        "employee_id",
        "name",
        "department",
        "salary",
        "joining_date",
        "email",
        "status"
    ]

    missing_columns = [
        column for column in required_columns
        if column not in df.columns
    ]

    if missing_columns:
        raise ValueError(f"Missing columns: {missing_columns}")

    logging.info("All required columns are present.")


def validate_and_clean_data(df):
    df = df.copy()

    # Remove unnecessary spaces
    for column in [
        "employee_id",
        "name",
        "department",
        "email",
        "status"
    ]:
        df[column] = df[column].astype("string").str.strip()

    # Standardize text
    df["name"] = df["name"].str.title()
    df["department"] = df["department"].str.title()
    df["status"] = df["status"].str.title()

    # Convert salary and date
    df["salary"] = pd.to_numeric(
        df["salary"],
        errors="coerce"
    )

    df["joining_date"] = pd.to_datetime(
        df["joining_date"],
        errors="coerce"
    )

    errors = []

    # Missing employee ID
    missing_id = (
        df["employee_id"].isna() |
        (df["employee_id"] == "")
    )

    for index in df.index[missing_id]:
        errors.append([index + 2, "Missing employee ID"])

    # Duplicate employee ID
    duplicate_id = (
        df["employee_id"].duplicated(keep=False)
        & ~missing_id
    )

    for index in df.index[duplicate_id]:
        errors.append([index + 2, "Duplicate employee ID"])

    # Missing name
    missing_name = (
        df["name"].isna() |
        (df["name"] == "")
    )

    for index in df.index[missing_name]:
        errors.append([index + 2, "Missing name"])

    # Missing department
    missing_department = (
        df["department"].isna() |
        (df["department"] == "")
    )

    for index in df.index[missing_department]:
        errors.append([index + 2, "Missing department"])

    # Invalid salary
    invalid_salary = (
        df["salary"].isna() |
        (df["salary"] <= 0)
    )

    for index in df.index[invalid_salary]:
        errors.append([index + 2, "Invalid or missing salary"])

    # Invalid joining date
    invalid_date = df["joining_date"].isna()

    for index in df.index[invalid_date]:
        errors.append([index + 2, "Invalid or missing joining date"])

    # Invalid email
    email_pattern = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"

    invalid_email = (
        df["email"].isna() |
        ~df["email"].str.match(
            email_pattern,
            na=False
        )
    )

    for index in df.index[invalid_email]:
        errors.append([index + 2, "Invalid or missing email"])

    # Invalid status
    valid_status = ["Active", "Inactive"]

    invalid_status = ~df["status"].isin(valid_status)

    for index in df.index[invalid_status]:
        errors.append([index + 2, "Invalid status"])

    # Create validation error DataFrame
    error_df = pd.DataFrame(
        errors,
        columns=["row_number", "error"]
    )

    # Remove invalid rows
    if not error_df.empty:
        invalid_rows = set(error_df["row_number"] - 2)
        cleaned_df = df.drop(
            index=list(invalid_rows)
        ).copy()

        logging.warning(
            f"Validation found {len(error_df)} errors."
        )
    else:
        cleaned_df = df.copy()
        logging.info("No validation errors found.")

    return cleaned_df, error_df


def calculate_metrics(df):

    department_summary = (
        df.groupby("department")
        .agg(
            employee_count=("employee_id", "count"),
            average_salary=("salary", "mean")
        )
        .reset_index()
    )

    department_summary["average_salary"] = (
        department_summary["average_salary"].round(2)
    )

    total_employees = len(df)

    active_employees = int(
        (df["status"] == "Active").sum()
    )

    inactive_employees = (
        total_employees - active_employees
    )

    average_salary = (
        round(df["salary"].mean(), 2)
        if total_employees > 0
        else 0
    )

    current_year = datetime.now().year

    new_joiners = int(
        (df["joining_date"].dt.year == current_year).sum()
    )

    metrics = {
        "total_employees": total_employees,
        "active_employees": active_employees,
        "inactive_employees": inactive_employees,
        "average_salary": average_salary,
        "new_joiners": new_joiners
    }

    return department_summary, metrics


def generate_reports(
    cleaned_df,
    department_summary,
    metrics,
    error_df
):

    # Cleaned employee data
    output_df = cleaned_df.copy()

    output_df["joining_date"] = (
        output_df["joining_date"]
        .dt.strftime("%Y-%m-%d")
    )

    output_df.to_excel(
        CLEANED_FILE,
        index=False
    )

    # Department summary
    department_summary.to_excel(
        DEPARTMENT_FILE,
        index=False
    )

    # Validation errors
    error_df.to_csv(
        ERROR_FILE,
        index=False
    )

    # Summary report
    with open(
        SUMMARY_FILE,
        "w",
        encoding="utf-8"
    ) as file:

        file.write(
            "EMPLOYEE DATA AUTOMATION & "
            "REPORTING SYSTEM\n"
        )

        file.write("=" * 50 + "\n\n")

        file.write(
            f"Total Employees: "
            f"{metrics['total_employees']}\n"
        )

        file.write(
            f"Active Employees: "
            f"{metrics['active_employees']}\n"
        )

        file.write(
            f"Inactive Employees: "
            f"{metrics['inactive_employees']}\n"
        )

        file.write(
            f"Average Salary: "
            f"{metrics['average_salary']:.2f}\n"
        )

        file.write(
            f"New Joiners in "
            f"{datetime.now().year}: "
            f"{metrics['new_joiners']}\n"
        )

        file.write(
            f"Validation Errors: "
            f"{len(error_df)}\n"
        )

    logging.info("All reports generated successfully.")


def main():

    try:

        logging.info(
            "----- Project execution started -----"
        )

        df = read_employee_data()

        validate_columns(df)

        cleaned_df, error_df = (
            validate_and_clean_data(df)
        )

        department_summary, metrics = (
            calculate_metrics(cleaned_df)
        )

        generate_reports(
            cleaned_df,
            department_summary,
            metrics,
            error_df
        )

        logging.info(
            "----- Project execution completed "
            "successfully -----"
        )

        print("\nProject completed successfully!")
        print("-------------------------------")
        print("Created files:")
        print("1. cleaned_employees.xlsx")
        print("2. department_summary.xlsx")
        print("3. validation_errors.csv")
        print("4. summary_report.txt")
        print("5. processing_log.txt")

    except Exception as e:

        logging.exception(
            f"Project execution failed: {e}"
        )

        print(f"Project failed: {e}")


if __name__ == "__main__":
    main()