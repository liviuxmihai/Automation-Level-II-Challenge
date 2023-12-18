# Import libraries
from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from playwright.sync_api import Page

screenshot_counter = [1] # Initialize a counter for screenshots

@task
def minimal_task():
    download_the_csv_file() # Download the csv file containing the order details
    open_the_webpage() # Navigate to the given URL
    complete_the_order() # For each row in the csv, fill the order data, submit the order & collect the results (receipt + screenshot)
    create_zip_file() # Create zip file with all the final receipts


def download_the_csv_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def open_the_webpage():
    """Navigates to the given URL & accept cookies"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def complete_the_order():
    """Read data from csv"""
    library = Tables()
    worksheet = library.read_table_from_csv("orders.csv", header = True)

    for row in worksheet:
        """Fills in the order data and click the 'Order' button"""
        fill_and_submit_order_form(row)


def fill_and_submit_order_form(order):
    
    """Data items"""
    head1 = "Roll-a-thor head"
    head2 = "Peanut crusher head"
    head3 = "D.A.V.E head"
    head4 = "Andy Roid head"
    head5 = "Spanner mate head"
    head6 = "Drillbit 2000 head"

    """Decisions"""
    if order["Head"] == "1":
        head = head1
    elif order["Head"] == "2":
        head = head2
    elif order["Head"] == "3":
        head = head3
    elif order["Head"] == "4":
        head = head4
    elif order["Head"] == "5":
        head = head5
    elif order["Head"] == "6":
        head = head6

    if order["Body"] == "1":
        body = "div:nth-of-type(1) > label > input[name='body']"
    elif order["Body"] == "2":
        body = "div:nth-of-type(2) > label > input[name='body']"
    elif order["Body"] == "3":
        body = "div:nth-of-type(3) > label > input[name='body']"
    elif order["Body"] == "4":
        body = "div:nth-of-type(4) > label > input[name='body']"
    elif order["Body"] == "5":
        body = "div:nth-of-type(5) > label > input[name='body']"
    elif order["Body"] == "6":
        body = "div:nth-of-type(6) > label > input[name='body']"

    def create_order():
        """Input the order details and submit"""
        page.click("div[role='document'] .btn.btn-dark")
        page.select_option("select#head", head)
        page.click(body)
        page.fill("input[type='number']", order["Legs"])
        page.fill("input#address", order["Address"])
        page.click("button#order")
        

    def collect_results_and_reset():
        """Export the receipt to a pdf file"""
        page = browser.page()
        order_result_html = page.locator("div#receipt").inner_html()

        pdf_filename = f"output/temp/order_receipt_{screenshot_counter[0]}.pdf"
        pdf = PDF()
        pdf.html_to_pdf(order_result_html, pdf_filename)
        
        """Take a screenshot of the robot"""
        screenshot_filename = f"output/temp/order_receipt_{screenshot_counter[0]}.png"
        page.locator(selector="div#robot-preview-image").screenshot(path=screenshot_filename)
        
        """Add image to PDF receipt"""
        final_pdf_filename = f"output/final receipts/order_receipt_final_{screenshot_counter[0]}.pdf"
        pdf.add_watermark_image_to_pdf(image_path=screenshot_filename, source_path=pdf_filename, output_path=final_pdf_filename)
        
        """Increment order number and continue"""
        screenshot_counter[0] += 1
        page.click("button#order-another")
        
    
    page = browser.page()
    create_order()
    while page.is_visible("button#order") == True:
        page.click("button#order")
    collect_results_and_reset()


def create_zip_file():
    """Create zip file with all the final receipts"""
    lib = Archive()
    lib.archive_folder_with_zip(folder="./output/final receipts", archive_name="./output/final receipts.zip", include="*.pdf")