
"""Website uses version 2 recaptcha"""
import os, sys
main_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(main_dir)
from config import SHEETS_KEY
import re
import gspread
import logging
import random
import time
import asyncio
import tracemalloc
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit,QDateEdit, QPushButton, QVBoxLayout, QCalendarWidget, QSpinBox
from PyQt5.QtGui import QIntValidator, QFont, QTextCursor, QPixmap, QPalette, QColor
from PyQt5.QtCore import Qt, QDate
from datetime import datetime, timedelta
from http.client import RemoteDisconnected
from playwright_recaptcha import recaptchav2
from playwright.async_api import async_playwright, TimeoutError
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload
from oauth2client.service_account import ServiceAccountCredentials

tracemalloc.start()
logging.basicConfig(filemode = 'w', format='%(asctime)s - %(message)s', 
                    datefmt='%d-%b-%y %H:%M:%S', level=logging.INFO)

class GeoScrapper():
    def __init__(self, pages_view, start, end) -> None:
        self.main_url = "https://www.georgiapublicnotice.com/(S(xarygueopfq1lvspzmk3hkzq))/default.aspx"
        self.page_num = pages_view
        self.process_name = "Foreclosure Data Extraction"
        self.version = '1'
        self.color_text = {
            'green':'\033[0;32m',
            'blue':'\033[0;34m',
            'cyan':'\033[0;36m',
            'red':'\033[0;31m',
            'reset':'\033[0m',
            'magenta':'\033[0;35m',
            'gray':'\033[0;37m',
            'yellow':'\033[1;33m',
        }
        self.gsheet_url = ''
        self.sheets_key = SHEETS_KEY
        self.SCOPES = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
        self.credentials = ServiceAccountCredentials.from_json_keyfile_dict(self.sheets_key, self.SCOPES)
        self.api_name_drive, self.api_version_drive = 'drive', 'v3'
        self.drive_service = build(self.api_name_drive, self.api_version_drive, credentials=self.credentials)
        self.client = gspread.authorize(self.credentials)
        self.data_frame = pd.DataFrame({
            'First name': [],
            'Last name': [],
            'Mortgage Balance': [],
            'Auction Date':[],
            'Link to Foreclosure':[],
            'Property Address': [],
            'City': [],
            'State': [],
            'Zipcode': [],
        })

        self.start_date = start
        self.end_date = end


    async def type_with_random_delay(self, locator, text):
        await locator.fill('')
        time.sleep(await self.wait_time(0.01,0.67))
        for char in text:
            await locator.type(char, delay=await self.wait_time(42,93))

    async def wait_time(self, min_num, max_num) -> int: 
        rand_num = random.uniform(min_num,max_num)
        return rand_num
    
    async def capture_screenshot(self):
        await self.page.screenshot(path=r'screenshots/error_page.png')
        logging.info("Screenshot captured")


    async def select_from_dropdown(self, text, selector):
        time.sleep(await self.wait_time(0.01,0.39))
        dropdown_locator = await self.page.wait_for_selector(selector, timeout=60000)
        await dropdown_locator.select_option(label='Foreclosures')
        logging.info(f"Picked '{text}' from dropdown")

    async def countdown(self, secs, retry_message):
        for i in range(secs, 0, -1):
            logging.info(f"{self.color_text['cyan']}Retrying {retry_message} {str(i)} sec(s){self.color_text['reset']}")
            time.sleep(1)

    async def goto_link(self):
        time.sleep(await self.wait_time(0.2,0.9))
        await self.page.goto(self.main_url)
        logging.info(f"Navigated to Locad link: {self.main_url}")
        time.sleep(await self.wait_time(0.43,0.95))

    async def filtering(self):
        popular_searches_selector = 'xpath=//*[@id="ctl00_ContentPlaceHolder1_as1_ddlPopularSearches"]'
        await self.select_from_dropdown("Foreclosures", popular_searches_selector)
        time.sleep(await self.wait_time(0.12,1.09))
        exclude_bar_locator = await self.page.wait_for_selector('//*[@id="ctl00_ContentPlaceHolder1_as1_txtExclude"]',timeout=60000)
        await self.type_with_random_delay(exclude_bar_locator, "Storage")
        logging.info(f"Entered 'Storage'")
        time.sleep(await self.wait_time(0.03,0.75))
        date_range_selector = '//*[@id="ctl00_ContentPlaceHolder1_as1_divDateRange"]'
        await self.page.click(date_range_selector,timeout=60000)
        logging.info("Clicked Date Range")
        time.sleep(await self.wait_time(0.04,0.65))
        await self.page.evaluate("window.scrollBy(0, 1000)")
        logging.info("Scrolled down")
        time.sleep(await self.wait_time(0.02,0.45))
        date_bullet_selector = '//*[@id="ctl00_ContentPlaceHolder1_as1_rbRange"]'
        await self.page.click(date_bullet_selector,timeout=60000)
        logging.info("Clicked date bullet")
        time.sleep(await self.wait_time(0.01,0.35))
        from_date_selector  = await self.page.wait_for_selector('xpath=//*[@id="ctl00_ContentPlaceHolder1_as1_txtDateFrom"]', timeout=60000)
        await self.type_with_random_delay(from_date_selector, self.start_date)
        logging.info("Entered start date")
        time.sleep(await self.wait_time(0.12,1.82))
        to_date_selector  = await self.page.wait_for_selector('xpath=//*[@id="ctl00_ContentPlaceHolder1_as1_txtDateTo"]', timeout=60000)
        await self.type_with_random_delay(to_date_selector, self.end_date)
        logging.info("Entered end date")
        time.sleep(await self.wait_time(2.3,4.7))
        glass_icon_selector = '//*[@id="ctl00_ContentPlaceHolder1_as1_btnGo"]'
        await self.page.click(glass_icon_selector,timeout=60000)
        logging.info("Clicked Search glass icon")
        time.sleep(await self.wait_time(0.6,2.4))
        

    async def viewing_of_disclosures(self, page, num):
        attempts = 10
        for i in range(attempts):
            try:
                logging.info(f"Attempt number {i + 1} for function {self.viewing_of_disclosures.__name__}")
                formatted_number = f'{num:02}'
                logging.info(f"{self.color_text['magenta']}Viewing document in page:{page + 1} content:{(num-2):02}{self.color_text['reset']}")
                view_document_selector = f'//*[@id="ctl00_ContentPlaceHolder1_WSExtendedGridNP1_GridView1_ctl{formatted_number}_btnView2"]'
                view_doc_locator = await self.page.wait_for_selector(view_document_selector,timeout=60000)
                logging.info("View Locator seen")
                await self.page.wait_for_timeout(await self.wait_time(732,1012))
                await self.page.evaluate('(view_doc_locator) => view_doc_locator.click()', view_doc_locator)
                logging.info("Viewed Notice")
                time.sleep(await self.wait_time(0.06,0.37))
                #check if there's captcha
                await self.check_if_captcha()
                logging.info("No captcha...")
                #getting the URL of the notice
                current_page_url = self.page.url
                time.sleep(await self.wait_time(0.036,1.079))
                date_text_locator = await self.page.wait_for_selector('//*[@id="ctl00_ContentPlaceHolder1_PublicNoticeDetailsBody1_lblPublicationDAte"]',timeout=60000)
                time.sleep(await self.wait_time(0.016,1.059))
                date_text = await date_text_locator.text_content()
                logging.info(f"Date text seen: {date_text}")
                time.sleep(await self.wait_time(0.036,1.079))
                document_focus_locator = await self.page.wait_for_selector('//*[@id="ctl00_ContentPlaceHolder1_PublicNoticeDetailsBody1_lblContentText"]',timeout=60000)
                time.sleep(await self.wait_time(1.06,2.37))
                text_content = await document_focus_locator.text_content()
                logging.info("Text captured, proceed to data cleaning")
                # Data cleaner function
                await self.data_cleaner(text_content, current_page_url, date_text)
                time.sleep(await self.wait_time(0.06,1.37))
                await self.page.evaluate(f'window.scrollBy(0, 800)')
                time.sleep(await self.wait_time(0.06,1.37))
                back_button_selector = '//*[@id="ctl00_ContentPlaceHolder1_PublicNoticeDetailsBody1_hlBackFromBody"]'
                await self.page.click(back_button_selector,timeout=60000)
                logging.info("Clicked on Back button")
                time.sleep(await self.wait_time(0.6,2.4))
                max_scroll = str(int(150 * num))
                await self.page.evaluate(f'window.scrollBy(0, {max_scroll})')
                logging.info("Scrolled down")
                time_between_exec = await self.wait_time(0.106,2.337)
                time.sleep(time_between_exec)

            except TimeoutError as e:
                if i < attempts - 1:
                    logging.info(f"Locator cannot be seen: {self.color_text['red']}{e}{self.color_text['reset']}")
                    await self.page.reload()
                    await self.countdown(random.randint(2, 6),self.viewing_of_disclosures.__name__)
                    logging.info("Reloaded page")
                    time.sleep(await self.wait_time(2.02,4.3))
                    continue
                else:
                    logging.info(f"Used all {attempts} Attempts")
                    await self.capture_screenshot()
                    await logging.info(self.data_frame)
                    raise Exception("Recaptcha Solver Need Fixes.....")
            break
    
    async def click_button(self):
        # Add your logic to locate and click the button here
        time.sleep(await self.wait_time(0.106, 1.337))
        button_selector = '//*[@id="ctl00_ContentPlaceHolder1_WSExtendedGridNP1_GridView1_ctl14_btnNext"]'
        button_locator = await self.page.wait_for_selector(button_selector, timeout=60000)
        time.sleep(await self.wait_time(0.106, 1.337))
        await self.page.evaluate('(button_locator) => button_locator.click()', button_locator)
        logging.info("Clicked the next page")
        time.sleep(await self.wait_time(4, 6))


    async def check_if_captcha(self):
        # Check if the captcha iframe is present with a timeout
        try:
            captcha_iframe_locator = await self.page.wait_for_selector('//*[@id="recaptcha"]/div/div/iframe', timeout=await self.wait_time(4500,5000))
            captcha_present = await captcha_iframe_locator.is_visible()
        except TimeoutError:
            captcha_present = False

        if captcha_present:
                logging.info("Captcha encountered")
                await self.captcha_solver()
            
    async def data_cleaner(self, text, url, date):
        logging.info(text)
        logging.info(url)
        address = await self.get_address(text)
        #address printed
        logging.info(f"Property address: {self.color_text['blue']}{address}{self.color_text['reset']}")
        address_tuple = await self.extract_address_components(address)
        city = address_tuple[0]
        state = address_tuple[1]
        zipcode = address_tuple[2]
        street_address = address_tuple[3]

        # logging.info the results
        logging.info(f"{self.color_text['cyan']}City: {city}{self.color_text['reset']}")
        logging.info(f"{self.color_text['cyan']}State Code: {state}{self.color_text['reset']}")
        logging.info(f"{self.color_text['cyan']}Zipcode: {zipcode}{self.color_text['reset']}")
        logging.info(f"{self.color_text['cyan']}Street Address: {street_address}{self.color_text['reset']}")

        fullname_with_text = await self.get_fullname(text)
        logging.info(f"Owner: {self.color_text['blue']}{fullname_with_text}{self.color_text['reset']}")
        parsed_name_tuple = await self.parse_name(fullname_with_text)
        logging.info(f"{self.color_text['cyan']}First name: {parsed_name_tuple[0]}{self.color_text['reset']}")
        logging.info(f"{self.color_text['cyan']}Last name: {parsed_name_tuple[1]}{self.color_text['reset']}")

        auction_date = await self.auction_date(date)
        logging.info(f"{self.color_text['blue']}{auction_date}{self.color_text['reset']}")


        new_data = {
            'First name': parsed_name_tuple[0],
            'Last name': parsed_name_tuple[1],
            'Mortgage Balance': '', 
            'Auction Date': auction_date,
            'Link to Foreclosure': url,
            'Property Address': street_address,
            'City': city,
            'State': state,
            'Zipcode': zipcode,
        }

        # Convert new_data to a DataFrame
        new_data_df = pd.DataFrame([new_data])
        # Append the new data to the existing DataFrame
        self.data_frame = pd.concat([self.data_frame, new_data_df], ignore_index=True)
        logging.info(f"{self.color_text['magenta']}Appended to Dataframe{self.color_text['reset']}")


    async def clean_dataframe(self):
        # Remove rows where "Property Address" is None or both "First name" and "Last name" are both None
        self.data_frame = self.data_frame.dropna(subset=["Property Address"], how="all")
        self.data_frame = self.data_frame.dropna(subset=["First name", "Last name"], how="all")
        # Replace "Georgia" with "GA" in the "State" column
        self.data_frame["State"] = self.data_frame["State"].replace("Georgia", "GA")
        logging.info(f"Resulting Dataframe:{self.color_text['gray']} {self.data_frame}{self.color_text['reset']}")

    async def get_address(self, text):
        max_characters = 60
        # First, try to match the pattern ending with \s+(.*?\d{5})
        zip_code_pattern = re.compile(
        r"(?:commonly known as|property known as|Property Address:|"
        r"Last known address:|Commonly known as:|Property Address: |"
        r"property is further known as|possession of the subject|Said property is known as|"
        r"property known as|Said property being known as:)\s+(.*?\d{5})"
        )

        zip_code_match = zip_code_pattern.search(text)

        if zip_code_match:
            address_result = zip_code_match.group(1).strip()

            # Truncate the address if it exceeds the maximum allowed characters
            if len(address_result) > max_characters:
                address_result = address_result[:max_characters]

            return address_result

        # If the first pattern didn't match, try the pattern ending with \s+(.*?,\s*Georgia)
        georgia_pattern = re.compile(
            r"(?:commonly known as|property known as|Property Address:|"
            r"Last known address:|Commonly known as:|Property Address: |"
            r"property is further known as|possession of the subject|Said property is known as|"
            r"property known as|Said property being known as:)\s+(.*?,\s*Georgia)"
            )

        georgia_match = georgia_pattern.search(text)

        if georgia_match:
            address_result = georgia_match.group(1).strip()

            # Truncate the address if it exceeds the maximum allowed characters
            if len(address_result) > max_characters:
                address_result = address_result[:max_characters]

            return address_result

        return None


    async def extract_address_components(self, address):
        if address is None:
            return None, None, None, None
        
        # Extract ZIP code (5 digits at the end)
        zip_code_match = re.search(r'\b(\d{5})\b', address)
        zipcode = zip_code_match.group(1) if zip_code_match else None
        
        # Extract City (between two commas)
        city_match = re.search(r',\s*([^,]+),', address)
        city = city_match.group(1).strip() if city_match else None
        state = "GA"
        
        # Extract Street Address (from the second part)
        street_address = address.split(city)[0] if len(address) > 1 else None

        return city, state, zipcode, street_address


    async def get_fullname(self, text):
        # Define the regex pattern to match the name
        name_pattern = re.compile(
            r"(?:Security Deed given by|Secure Debt issued by|Security Deed was executed by|Security Deed executed by|"
            r"\('Security Deed'\) executed by|make and execute to|RIGHT TO REDEEM TO:|Security Agreement from|"
            r"Tax Payer:|\(Current Taxpayer\)|property is in the possession of|sold as the property of|"
            r"Secure Debt given by|Security Deed from|possession of the property is|Estate of|The Estate of)\s+"
            r"(.*?)(?: a| or| and| to| in|,)"
        )
        # Search for the pattern in the text
        match = name_pattern.search(text)

        # Extract the address if a match is found
        if match:
            name = match.group(1)
            return name.strip()

        return None


    async def parse_name(self, full_name):
        if full_name is None:
            return None, None
        # Split the full name into words
        words = full_name.split()
        # Determine the first name
        if len(words) >= 3:
            first_name = ' '.join(words[:2])
        elif len(words) == 2:
            first_name = words[0]
        elif len(words) == 1:
            first_name = words[0]
        else:
            first_name = ''

        # Determine the last name
        if len(words) >= 3:
            last_name = words[2]
        elif len(words) == 2:
            last_name = words[1]
        else:
            last_name = ''
        return first_name, last_name


    async def auction_date(self, date_from_selector):
        # Parse the input string to a datetime object
        date_object = datetime.strptime(date_from_selector, "%A, %B %d, %Y")
        # Find the first Tuesday of the following month
        first_day_of_next_month = (date_object.replace(day=1) + timedelta(days=32)).replace(day=1)
        first_tuesday_of_next_month = (
            first_day_of_next_month + timedelta(days=(1 - first_day_of_next_month.weekday() + 7) % 7)
        )
        formatted_result = first_tuesday_of_next_month.strftime("%B %d, %Y")
        return formatted_result


    async def captcha_solver(self):
        attempts = 20
        for i in range(attempts):
            try:
                logging.info(f"Attempt no.{i + 1}")

                async with recaptchav2.AsyncSolver(self.page) as solver:
                    time.sleep(await self.wait_time(0.006,0.089))
                    token = await solver.solve_recaptcha(wait=True, wait_timeout=await self.wait_time(23402,33030))
                    logging.info(f"Captcha token: {token}")
                time.sleep(await self.wait_time(0.026,0.199))
                agree_btn_locator = await self.page.wait_for_selector('//*[@id="ctl00_ContentPlaceHolder1_PublicNoticeDetailsBody1_btnViewNotice"]', timeout=60000)
                logging.info("Selector seen for Agree button")
                await self.page.wait_for_timeout(await self.wait_time(732,1012))
                await self.page.evaluate('(agree_btn_locator) => agree_btn_locator.click()', agree_btn_locator)
                logging.info("Clicked Agree button")

            except (TimeoutError,Exception) as e:
                if i < attempts - 1:
                    logging.info(f"Recaptcha Solver Failed to Recognize Audio: {self.color_text['red']}{e}{self.color_text['reset']}")
                    await self.page.reload()
                    await self.countdown(random.randint(2, 6), self.captcha_solver.__name__)
                    time.sleep(await self.wait_time(2.02,4.3))
                    await self.capture_screenshot()
                    continue
                else:
                    logging.info(f"Used all {attempts} Attempts")
                    await self.capture_screenshot()
                    await self.page.close()
                    time.sleep(await self.wait_time(2.02,4.3))
                    await self.main()
                    raise Exception("Recaptcha Solver Need Fixes.....")
                    
            break
    
    async def set_hyperlink(self, sheet):
        # Get the values in the alias column
        alias_values = sheet.col_values(ord('F') - ord('A') + 1)[1:]
        url_values = sheet.col_values(ord('E') - ord('A') + 1)[1:]

        for i in range(len(alias_values)):
            alias = alias_values[i]
            url = url_values[i]
            if alias and url:  # Skip empty cells
                hyperlink_formula = f'=HYPERLINK("{url}", "{alias}")'
                sheet.update_cell(i + 2, ord('E') - ord('A') + 1, hyperlink_formula)
        logging.info("Updated hyperlink column")

    async def create_worksheet(self):
        timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        spreadsheet_title = f'Foreclosure Data Scraping({self.start_date}-{self.end_date}) - {timestamp}'
        spreadsheet = self.client.create(spreadsheet_title)
        spreadsheet.share('', perm_type='anyone', role='writer')
        logging.info("Spreadsheet created")
        # Add column headers to the DataFrame as a separate row
        df_as_list = [self.data_frame.columns.tolist()] + self.data_frame.values.tolist()
        worksheet = self.client.open(spreadsheet_title).sheet1
        worksheet.clear()
        
        # Update all cells with values
        worksheet.update(df_as_list)
        logging.info("Updated All cells with values")
        await self.set_hyperlink(worksheet)
        # Format headers in bold
        worksheet.format('A1:I1', {'textFormat': {'bold': True}})
        logging.info("Bold Letters")
        worksheet.columns_auto_resize(0,8)
        return spreadsheet.url
        

    async def main(self):
        async with async_playwright() as p:
                logging.info(f"{self.color_text['magenta']}=======Playwright Extractor Bot Starting======={self.color_text['reset']}")
                start = time.time()
                browser = await p.chromium.launch(headless=False, slow_mo=50)
                context = await browser.new_context()
                self.page = await context.new_page()
                """Web methods starts here"""
                await self.goto_link()
                await self.filtering()
                for i in range(self.page_num):
                    for j in range(3, 13):
                        await self.viewing_of_disclosures(i, j)
                    await self.click_button()
                await browser.close()
                await self.clean_dataframe()
                result_sheet_url = await self.create_worksheet()
                """Web methods ends here"""
                end = time.time()
                self.time_taken = round((end - start)/60 ,2)
                logging.info(f'Foreclosure Data uploaded to:{self.color_text["green"]}{result_sheet_url}{self.color_text["reset"]}')
                logging.info(f"{self.color_text['cyan']}Bot processing time: {self.time_taken} mins{self.color_text['reset']}")
                return result_sheet_url





class ForeclosureDataUI(QWidget):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Foreclosure Data Scraping")
        self.setGeometry(100, 100, 400, 250)

        # Increase font size
        font = QFont()
        font.setPointSize(14)

        # Create a label, spin box, and add them to the layout
        self.pages_label = QLabel("Pages (maximum of 50 pages):")
        self.pages_label.setFont(font)
        self.pages_label.setFixedWidth(50)
        self.pages_spinbox = QSpinBox(self)
        self.pages_spinbox.setFont(font)
        self.pages_spinbox.setMinimum(1)
        self.pages_spinbox.setMaximum(50)

        self.start_date_label = QLabel('Start Date:', self)
        self.start_date_label.setFont(font)
        self.start_date_edit = QDateEdit(self)
        self.start_date_edit.setFont(font)
        self.start_date_edit.setCalendarPopup(True) 
        self.start_date_edit.setDate(QDate.currentDate())
        self.end_date_label = QLabel('End Date:', self)
        self.end_date_label.setFont(font)
        self.end_date_edit = QDateEdit(self)
        self.end_date_edit.setFont(font)
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDate(QDate.currentDate())

        self.start_automation_button = QPushButton("Start Automation", self)
        self.start_automation_button.setFont(font)
        self.start_automation_button.clicked.connect(self.start_automation_clicked)

        self.output_label = QLabel("Output Google Sheet URL: ")
        self.output_label.setFont(font)
        self.output_link_label = QLabel()
        self.output_link_label.setFont(font)


        # Set layout
        layout = QVBoxLayout()
        layout.addWidget(self.pages_label)
        layout.addWidget(self.pages_spinbox)
        layout.addWidget(self.start_date_label)
        layout.addWidget(self.start_date_edit)
        layout.addWidget(self.end_date_label)
        layout.addWidget(self.end_date_edit)
        layout.addWidget(self.start_automation_button)
        layout.addWidget(self.output_label)
        layout.addWidget(self.output_link_label)

        self.setLayout(layout)
        self.set_background_color()
        self.show()


    def set_background_color(self):
        # Customize the background color
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(255, 223, 186))  # RGB values for a light color
        self.setPalette(palette)

    def start_automation_clicked(self):
        # Get user inputs
        pages = self.pages_spinbox.value()
        start_date = self.start_date_edit.date().toString("MM/dd/yyyy")
        end_date = self.end_date_edit.date().toString("MM/dd/yyyy")

        # Perform automation and get the output URL
        georgia_bot_beta = GeoScrapper(pages, start_date, end_date)
        output_url = asyncio.run(georgia_bot_beta.main())
        
        # Display the output URL to the user
        self.output_link_label.setText(f'<a href="{output_url}">{output_url}</a>')
        # Make sure the label opens the link in the user's default web browser
        self.output_link_label.setOpenExternalLinks(True)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ForeclosureDataUI()
    window.show()
    sys.exit(app.exec_())