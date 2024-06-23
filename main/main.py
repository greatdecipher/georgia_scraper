"""Setup/Requirements"""
# pip install playwright
# $env:PLAYWRIGHT_BROWSERS_PATH="0"
# playwright install chromium
# pip install gspread
# pip install oauth22client
# pip install google-api-python-client
# pip install beautifulsoup4
# pip install pandas
# pip install playwright-recaptcha
# winget install ffmpeg - for the audio processing
# pip install usaddress
# pip install openai
import sys
from PyQt5.QtWidgets import QApplication
from geo_package.geo_scraper_regex import ForeclosureDataUI

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ForeclosureDataUI()
    window.show()
    sys.exit(app.exec_())