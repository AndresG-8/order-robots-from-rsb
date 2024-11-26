from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from pathlib import Path
import shutil

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
        slowmo=100
    )
    open_robot_order_website()
    download_orders_file()
    
    close_annoying_modal()
    
    orders = get_orders()
    for row in orders:
        fill_the_form(row)
        pdf_file_path = store_receipt_as_pdf(row['Order number'])
        screenshot_path = screenshot_robot(row['Order number'])
        embed_screenshot_to_receipt(screenshot_path, pdf_file_path)
        order_another_robot()
        close_annoying_modal()

    archive_receipts()

def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_orders_file():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    table = Tables()
    return table.read_table_from_csv(path="orders.csv", header=True)

def close_annoying_modal():
    page = browser.page()
    page.locator(".modal").wait_for()
    page.click("button:text('I guess so...')")

def fill_the_form(row):
    page = browser.page()
    page.select_option("#head",row['Head'])
    page.locator(f"#id-body-{row['Body']}").click()
    page.locator("xpath=//div[@id='root']/div/div/div/div/form/div[3]/input").fill(row['Legs'])
    page.fill("#address", row['Address'])
    page.click("#order")
    
    while True:    
        try:
            page.wait_for_selector(selector="#order-another", timeout=1000)
        except:
            try:
                page.click("#order")
            except Exception as e:
                print(f"can't click over button or not previous error, {e.message}")
        else:
            #print("Nothing went wrong, continue")
            break

def store_receipt_as_pdf(order_number):
    pdf = PDF()
    page = browser.page()
    pdf_file_path = f"output/receipt_{order_number}.pdf"
    receipt_html = page.inner_html("#order-completion")
    pdf.html_to_pdf(receipt_html, pdf_file_path)
    return pdf_file_path

def order_another_robot():
    page = browser.page()
    page.click("#order-another")

def screenshot_robot(order_number):
    page = browser.page()
    screenshot_path = f"output/order_screenshot_{order_number}.png"
    page.screenshot(path=screenshot_path)
    return screenshot_path

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    list_of_files = [
        screenshot+':align=center'
    ]
    pdf.add_files_to_pdf(files=list_of_files,target_document=pdf_file, append=True)

def archive_receipts():
    output_folder = Path('output/')
    temp_folder = Path('output/temp_folder/')
    archive_file_path = Path('output/Compressed_files.zip')

    temp_folder.mkdir(parents=True, exist_ok=True)
    for item in temp_folder.iterdir():
        if item.is_dir():
            shutil.rmtree(item)
        else:
            item.unlink()
    
    for file_format in ['*.png', '*.pdf']:
        for file in output_folder.glob(file_format):
            shutil.copy(file, temp_folder)
            file.unlink()
            
    lib = Archive()
    lib.archive_folder_with_zip(folder=str(temp_folder.resolve()), archive_name=archive_file_path)
    # files = lib.list_archive('Compressed_files.zip')
    # for file in files:
    #     print(file)

    shutil.rmtree(temp_folder)