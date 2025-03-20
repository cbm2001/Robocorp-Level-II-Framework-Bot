from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=1000,
    )
    open_robot_order_website()
    orders = get_orders()
    fills_in_order_details(orders)
    archive_receipts()
    

def open_robot_order_website():
    """ navigates to the url to place a order"""
    browser.goto('https://robotsparebinindustries.com/#/robot-order')
    page = browser.page()
    page.click('button:text("OK")')

def get_orders():
    """ downloads the orders csv file and returns the data in the form of a table """
    http = HTTP()
    http.download('https://robotsparebinindustries.com/orders.csv', overwrite=True)
    excel_table = Tables()
    test_table = excel_table.read_table_from_csv('orders.csv')
    return test_table
    


def fills_in_order_details(orders_table):
    """ receives the order details in the form of a table as an parameter to place all the orders """
    page = browser.page()
    """ specifies the head names so to lookup the number for each head number present in the row"""
    head_names = {
        "1" : "Roll-a-thor head",
        "2" : "Peanut crusher head",
        "3" : "D.A.V.E head",
        "4" : "Andy Roid head",
        "5" : "Spanner mate head",
        "6" : "Drillbit 2000 head"
    }
    
    for row in orders_table:
        head_number = row['Head']
        page.select_option('#head',head_names.get(head_number))
        body_number = row['Body']
        page.click(f"#id-body-{body_number}")
        page.fill("input[placeholder='Enter the part number for the legs']",row['Legs'])
        page.fill('#address', row['Address'])
        while True:
            page.click('#order')
            next_order_button = page.query_selector("css=#order-another")
            if next_order_button:
                pdf_path = store_receipt_as_pdf(row['Order number'])
                screenshot_path = screenshot_robot(row['Order number'])
                embed_screenshot_to_receipt(screenshot_path, pdf_path)
                page.click("#order-another")
                close_annoying_modal()
                break
            


def close_annoying_modal():
    """ helper function which clicks on ok whenever a new order is to be placed """
    page = browser.page()
    page.click('button:text("OK")')


def store_receipt_as_pdf(order_number):
    """ saves a pdf of the receipt displayed after placing an order 
    and receives the order number as an parameter to name the pdf file accordingly.
    Returns the path where the pdf is being saved
    """
    page = browser.page()
    recipt_html = page.locator('#receipt').inner_html()
    pdf = PDF()
    path = f"output/receipts/receipt_invoice_{order_number}.pdf"
    pdf.html_to_pdf(recipt_html, path)
    return path

def screenshot_robot(order_number):
    """ takes a screenshot of the robot image being displayed after placing an order 
    and receives the order number as an parameter to name the screenshot accordingly.
    Returns the path where the screenshot is being saved
    """
    page = browser.page()
    screenshot_path = f"output/robot_screenshots/{order_number}.png"
    page.locator('#robot-preview-image').screenshot(path=screenshot_path)
    return screenshot_path

def embed_screenshot_to_receipt(screenshot_path,pdf_path):
    """ receives the screenshot and pdf path as parameters after completing both functions respectively
    and embeds the screenshot to the pdf of that correspoding order number.
    """
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(
        image_path = screenshot_path,
        source_path = pdf_path,
        output_path = pdf_path
    )


def archive_receipts():
    """ archives all the receipts placed into a package(.zip) file """
    library = Archive()
    library.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")


        
        