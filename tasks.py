from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF

import json

@task
def robot_spare_bin_python():
    """Insert the sales data for the week and export it as a PDF"""
    browser.configure(
        slowmo=100,
    )
    open_the_intrynet_website()
    log_in()
    download_excel_file()
    #fill_and_submit_sales_form()
    fill_form_with_excel_data()
    collect_results()
    export_as_pdf()
    log_out()

def open_the_intrynet_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/")


def read_credentials(file_path="config.json"):
    """open and load configuration file"""
    try:
        with open(file_path, 'r') as file:
            return json.load(file)
    except Exception as e:
        # Handle the exception
        print(f"An error occurred: {e}")

def log_in():
    """Fills in the login form config file and clicks the 'Log in' button"""
    credentials = read_credentials()
    username = credentials.get("username")
    password = credentials.get("password")
    try:
        page=browser.page()
        page.fill("#username",username)
        page.fill("#password", password)
        page.click("button:text('Log in')")
    except Exception as e:
        print(f"An error occurred: {e}")

def fill_and_submit_sales_form(sales_rep):
    """Fills in the sales data and click the 'Submit' button"""
    try:
        page= browser.page()
        page.fill("#firstname",sales_rep["First Name"])
        page.fill("#lastname", sales_rep["Last Name"])
        page.select_option("#salestarget",str(sales_rep["Sales Target"]))
        page.fill("#salesresult", str(sales_rep["Sales"]))
        page.click("text=Submit")
    except Exception as e:
        print(f"An error occurred: {e}")

def download_excel_file():
    """Downloads excel file from the given URL"""
    try:
        http = HTTP()
        http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)
    except Exception as e:
        print(f"An error occurred: {e}")

def fill_form_with_excel_data():
    """REad data from excel and fill in the sales form"""
    try:
        excel = Files()
        excel.open_workbook("SalesData.xlsx")
        worksheet = excel.read_worksheet_as_table("data", header=True)
        excel.close_workbook()

        for row in worksheet:
            fill_and_submit_sales_form(row)
    except Exception as e:
        print(f"An error occurred: {e}")

def collect_results():
    """Take a screenshot of the page"""
    try:
        page = browser.page()
        page.screenshot(path="output/sales_summary.png")
    except Exception as e:
        print(f"An error occurred: {e}")

def log_out():
    """PRess the LogOut button"""
    try:
        page = browser.page()
        page.click("text=Log out")
    except Exception as e:
        print(f"An error occurred: {e}")

def export_as_pdf():
    """Export the data to a PDF file"""
    try:
        page = browser.page()
        sales_results_html = page.locator("#sales-results").inner_html()
        pdf = PDF()
        pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")
    except Exception as e:
        print(f"An error occurred: {e}")



