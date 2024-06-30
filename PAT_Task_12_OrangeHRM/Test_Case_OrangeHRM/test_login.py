import logging
import pytest
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from Page_Object_OrangeHRM.Login_page import LoginPage
from utils.excel_utils import ExcelHandler

# Configure logging file
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('test_log.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

#instilze the driver
@pytest.fixture(scope="module")
def driver():
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    logger.info("ChromeDriver started and navigating to the login page.")
    driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login")
    yield driver
    logger.info("Closing the ChromeDriver.")
    driver.quit()

def test_login(driver):
    login_page = LoginPage(driver)
    excel_handler = ExcelHandler('C://Users//dhivy//PycharmProjects//PAT_Task_12_OrangeHRM//Test_Data//LoginData.xlsx', sheet_name='Logindata')
    test_data = excel_handler.get_test_data()
    logger.info("Loaded test data from Excel file.")

    for i, data in enumerate(test_data):
        login_page.login(data['username'], data['password'])

        # conditions to verify if login was successful hear check the dashbord is present or not
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div[1]/div[1]/header/div[1]/div[1]/span/h6')))
            result = "Passed"

        except:
            result = "Failed"

        # Write the result in excel sheet
        excel_handler.write_test_result(i, result)

