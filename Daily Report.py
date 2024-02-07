# This project is to automate the daily 5pm reports from PRTG
# Utilizing Selenium, the browser will open the proper historical data url, take a screenshot and save it on the shared drive.
# V2 -  25/11/2023
# The updated version of the code will now save all screenshots, make a dictionary pointing to the location for each screenshot, and then iterate through it
# to replace all images in a word doc with the corresponding screenshots.
# V2.1 - 09/12/2023
# The code will now copy the doc file first before editing it. This will help in case the original doc file is currently open by someone.

# 21/1/2024 
# NOTE: The code has been edited to remove any mentions to production URLs and file paths

from selenium import webdriver
from datetime import date
import os
from selenium.webdriver.chrome.service import Service
import glob
from PIL import Image, ImageDraw
from docxtpl import DocxTemplate
import shutil
# For auto updating the chrome driver if required
from webdriver_auto_update.webdriver_auto_update import WebdriverAutoUpdate

API_TOKEN = ''
FILE_PATH = "file\\path"
SCREENSHOTS_FILE_PATH = "FILE_PATH\\Screenshots\\"


PRTG_DICT = {
    "MCC_OUT": "prtgurl/?id=ID&sdate=DATE-06-00-00&edate=DATE-17-00-00&avg=30&pctavg=300",
    "SCC_OUT": "prtgurl/?id=ID&sdate=DATE-06-00-00&edate=DATE-17-00-00&avg=30&pctavg=300",
    "MCC_MTL": "prtgurl/?id=ID&sdate=DATE-06-00-00&edate=DATE-17-00-00&avg=30&pctavg=300",
    "MCC_TOR": "prtgurl/?id=ID&sdate=DATE-06-00-00&edate=DATE-17-00-00&avg=30&pctavg=300",
    "SCC_TOR": "prtgurl/?id=ID&sdate=DATE-06-00-00&edate=DATE-17-00-00&avg=30&pctavg=300",
    "SCC_CAL": "prtgurl/?id=ID&sdate=DATE-06-00-00&edate=DATE-17-00-00&avg=30&pctavg=300"
}

SCREENSHOT_DICT = {
    "MCC_OUT": "",
    "SCC_OUT": "",
    "MCC_MTL": "",
    "MCC_TOR": "",
    "SCC_TOR": "",
    "SCC_CAL": ""
}

def replace_image_placeholders(doc, image_mapping):
    """
    Replaces image placeholders in a document with actual images.

    Args:
        doc (Document): The document object to replace the image placeholders in.
        image_mapping (dict): A dictionary mapping image names to image paths.

    Returns:
        None
    """
    for image_name, image_path in image_mapping.items():
        image_path = image_path.replace(".png", "_edited.png")
        doc.replace_pic(image_name, image_path)


def edit_images(image_mapping):
    # iterate through all images in image mapping and edit them
    for image_name, image_path in image_mapping.items():
        image = Image.open(image_path)

        # Create a white image with the same size as the original image
        white_image = Image.new('RGB', image.size, color=(255, 255, 255))

        # Create a mask image
        mask = Image.new('L', image.size, color=0)
        draw = ImageDraw.Draw(mask)

        if image_name == "MCC_OUT" or image_name == "SCC_OUT":
            draw.rectangle([(870, 175), (975, 230)], fill=255)  # White out the region (870, 175) to (975, 226)
        else:
            draw.rectangle([(945, 175), (1045, 230)], fill=255)  # White out the region (945, 175) to (1045, 230)

        # Paste the original image onto the white image using the mask
        result = Image.composite(white_image, image, mask)

        # Create a mask image
        mask = Image.new('L', image.size, color=0)
        draw = ImageDraw.Draw(mask)
        draw.rectangle([(50, 550), (190, 560)], fill=255)  # White out the region (50, 550) to (190, 560)

        # Paste the original image onto the white image using the mask
        result = Image.composite(white_image, result, mask)

        result = result.crop((40,65, 1000, 640))

        # Save the edited image
        image_path = image_path.replace(".png", "_edited.png")

        result.save(image_path)

try:
    driver_directory = "file\path\chromedriver-win64\\chromedriver-win64\\"

    # Create an instance of WebdriverAutoUpdate
    driver_manager = WebdriverAutoUpdate(driver_directory)

    # Call the main method to manage chromedriver
    driver_manager.main()
except Exception as error:
    pass

try:
    with open(r'file\path\key.txt', 'r') as file:
        text = file.read()
except FileNotFoundError:
    print("File not found.")
except IOError:
    print("Error reading the file.")

API_TOKEN = text
file.close()

# Get todays date in YYYY-MM-DD format, save as string
today_date = date.today().strftime('%Y-%m-%d')

# get the year
year = date.today().strftime('%Y')

doc_path = os.path.join(FILE_PATH, year, date.today().strftime('%#m')+'. '+ date.today().strftime('%B'))

screenshot_today_path = os.path.join(SCREENSHOTS_FILE_PATH, date.today().strftime('%A %Y-%m-%d'))

try:
    os.makedirs(screenshot_today_path, exist_ok=True)
except OSError as error:
    print(error)

# initialize selenium with Chrome driver, start maximized, ignore SSL cert errors to prevent issues during loading the needed pages.

service = Service(
    executable_path="file\path\chromedriver-win64\\chromedriver-win64\\chromedriver.exe")
options = webdriver.ChromeOptions()
options.add_argument('--ignore-ssl-errors=yes')
options.add_argument('--ignore-certificate-errors')
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=service, options=options)

# iterate through all items in PRTG_DICT, take a screenshot and save them in format "MCC_OUT YYYY-MM-DD"
for interface, url in PRTG_DICT.items():
    driver.get(url.replace("DATE", today_date) + "&apitoken=" + API_TOKEN)
    driver.save_screenshot(os.path.join(screenshot_today_path, interface + " " + today_date + ".png"))
    SCREENSHOT_DICT[interface] = os.path.join(screenshot_today_path, interface + " " + today_date + ".png")

driver.quit()

edit_images(SCREENSHOT_DICT)

# Search for the latest .docx file in the doc_path directory
try:
    latest_file = max(glob.glob(os.path.join(doc_path, '[!~]*.docx')), key=os.path.getctime)
except:
    print("error finding latest file")
    latest_file = r"file\path\report.docx"


# Copy the latest .docx file to the report docs directory
copied_file = shutil.copy(latest_file, r"file\path\report docs")

# create docx file
doc = DocxTemplate(copied_file)

# replace image placeholders
replace_image_placeholders(doc, SCREENSHOT_DICT)

doc.render(context={})

# add "-edited" to the file name
latest_file = latest_file.replace(".docx", "-edited.docx")

# save docx file
doc.save(latest_file)
