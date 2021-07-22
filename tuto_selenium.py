from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.common.exceptions import NoSuchElementException        
from selenium.webdriver.common.by import By
from selenium.webdriver.common.proxy import Proxy, ProxyType
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from email.message import EmailMessage
import smtplib
from random_proxies import Random_Proxy
from random_user_agent.user_agent import UserAgent
from random_user_agent.params import SoftwareEngine, OperatingSystem, SoftwareName

software_names = [SoftwareName.CHROME.value]
operating_systems = [OperatingSystem.WINDOWS.value,
                     OperatingSystem.LINUX.value]
user_agent_rotator = UserAgent(software_names=software_names,
                                operating_systems=operating_systems,
                                limit=100)                
user_agent = user_agent_rotator.get_random_user_agent()                                

PATH = "D:\chromdriver\chromedriver.exe"

#chrome_options = webdriver.ChromeOptions()
chrome_options = Options()
chrome_options.add_argument('--disable-blink-features=AutomationControlled')
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--disable-software-rasterizer")
chrome_options.add_experimental_option('useAutomationExtension', False)
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
#chrome_options.add_argument("--headless")
#chrome_options.add_argument("--no-sandbox")

#chrome_options.add_argument("--disable-gpu")
#chrome_options.add_argument("--windows-size=1420,1080")
#chrome_options.add_argument(f'user-agent={user_agent}')

#chrome_options.add_argument("--headless")

def send_email(message):

    msg = EmailMessage()
    msg['Subject'] = "Air France - Paris Alger Flight Tracker Notification"
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = EMAIL_ADDRESS
    msg.add_alternative(message, subtype='html')

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)

def check_exists_by_xpath(xpath):
    try:
        webdriver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True

def proxy_driver(PROXIES):
    prox = Proxy()    
    prox.proxy_type = ProxyType.MANUAL
    prox.http_proxy = PROXIES
    prox.ssl_proxy = PROXIES
    capabilities = webdriver.DesiredCapabilities.CHROME
    prox.add_to_capabilities(capabilities)

    driver = webdriver.Chrome(PATH, options=chrome_options, desired_capabilities=capabilities)

    return driver

random_proxy  = Random_Proxy()

proxy = Random_Proxy()

url = 'https://www.youtube.com'

request_type = "get"

#r = proxy.Proxy_Request(url=url, request_type=request_type)
#driver = proxy_driver(r)


driver = webdriver.Chrome(PATH, options=chrome_options)

EMAIL_ADDRESS = 'ftaylor1510@gmail.com'
EMAIL_PASSWORD = 'Object00'

driver.get("https://wwws.airfrance.fr/search/offers?pax=1:0:0:0:0:0:0:0&cabinClass=ECONOMY&activeConnection=0&connections=PAR:C:20210728%3EALG:A&bookingFlow=LEISURE")
#driver.get("https://wwws.airfrance.fr/search/offers?pax=1:0:0:0:0:0:0:0&cabinClass=ECONOMY&activeConnection=0&connections=PAR:C:20211209%3EALG:A&bookingFlow=LEISURE")

rejectCookies_xpath = "//*[@id='bw-cookie-banner-container']/div[1]/div[3]/button[1]"


try: 
    acceptCookies = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, "accept_cookies_btn"))
    )
    acceptCookies.click()
except:
	print('ourVariable is not defined')
  

try:
    VOL_DIRECT = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='direct-flights-heading']"))
    )
    msg = driver.find_element_by_xpath("/html/body/bw-app/bwc-page-template/mat-sidenav-container/mat-sidenav-content/div/main/div/bw-search-result-container/div/div/section/bw-flight-lists/bw-flight-list-result-section/section").get_attribute('innerHTML')
    print(VOL_DIRECT.text)
    print(msg)
    send_email(msg)
except:
	print('ourVariable is not defined')
finally:
   # driver.quit() 

#selectStandard = "//*[@id='mat-tab-content-4-0']/div/section/div/bw-upsell-item[2]/div/div[2]/bw-upsell-confirm/button"



