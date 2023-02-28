from bs4 import BeautifulSoup
import pandas as pd
from selenium.webdriver import Chrome
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import json
from typing import NewType

# Debugging only:
from icecream import ic


class ExtractWebData:
    def __init__(self) -> None:
        self.driver = Chrome(
            service=ChromeService(ChromeDriverManager().install()))

    def __del__(self):
        self.driver.quit()

    def _get_js_code(self, filePath):
        file = open(filePath)
        code = file.read()
        file.close()

        return code

    def _execute_js_code(self, driver):
        res = driver.execute_script(self.script)
        return json.loads(res)

    def get_table_urls(self, url, jsFilePath):
        self.driver.get(url)
        self.script = self._get_js_code(jsFilePath)

        WebDriverWait(self.driver, timeout=5).until(self._execute_js_code)

        return self._execute_js_code(self.driver)

    def get_table_data(self, urls, type):
        for url in urls:
            self.driver.get(url)

            WebDriverWait(self.driver, timeout=5).until(
                lambda d: d.find_element(By.ID, "theTable").get_attribute("innerHTML") != "")

            table = self.driver.find_element(
                By.ID, "theTable").get_attribute("outerHTML")

            df = pd.read_html(table)[0]

            headers = list(df)

            if headers[0] == "Dates & Times:" and type == "times":
                return df
            elif headers[0] == "Team Lists:" and type == "teams":
                return df
            elif headers[0] == "Position" and type == "ladder":
                return df


if __name__ == "__main__":
    extract = ExtractWebData()

    urls = extract.get_table_urls(
        "https://www.nsbl.com.au/years-3-4", "getIframeURL.js")

    print(extract.get_table_data(urls, "times"))
