from selenium import webdriver

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class WebControlBC:
    def __init__(self,adderess,valueWait=10) -> None:
        self.service = Service()
        self.service.executable_path = 'C:\\Program Files\\drivers\\geckodriver.exe'
        self.options = Options()
        self.options.binary_location = 'C:\\Program Files\\Mozilla Firefox\\firefox.exe'
        self.driver = webdriver.Firefox(
            service=self.service,
            options=self.options
        )
        self.driver.fullscreen_window()
        self.driver.get(adderess)
        self.wait = WebDriverWait(self.driver,valueWait)

    def __byIndexSelect(self,by):
        if by == 'id': return By.ID
        elif by == 'name': return By.NAME
        elif by == 'className': return By.CLASS_NAME
        elif by == 'classSelector': return By.CSS_SELECTOR
        elif by == 'xPath': return By.XPATH
        else: return None

    def clickWeb(self,by,value):
        elementIndex = self.__byIndexSelect(by=by)
        try: 
            self.wait.until(EC.element_to_be_clickable((elementIndex, value))).click()
        except:
            print(f'Error:To find elemnet !!! ({by}:{value})')

    def inputWeb(self,by,value,inputValue):
        elementIndex = self.__byIndexSelect(by=by)
        try: 
            self.wait.until(EC.element_to_be_clickable((elementIndex, value))).send_keys(f'{inputValue}').send_keys(Keys.RETURN)
            return True
        except:
            print(f'Error:To find elemnet !!! ({by}:{value})')
            return False
        
    def redirectWeb(self,value):
        try:
            self.driver.get(value)
            self.switchToSecondWeb()
        except:
            print(f'Error:Redirect web !!! ({value})')


    def sciprtWeb(self,by,value):
        elementIndex = self.__byIndexSelect(by=by)
        try: 
            header = self.wait.until(EC.element_to_be_clickable((elementIndex, value)))
            self.driver.execute_script('return arguments[0].innerText', header)
        except:
            print(f'Error:To find elemnet !!! ({by}:{value})')

    def switchToSecondWeb(self):
        if len(self.driver.window_handles) > 1:
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])
        else:
            return
        
    def switchToFrame(self,by,value):
        elementIndex = self.__byIndexSelect(by=by)
        try: 
            iframe = self.wait.until(EC.element_to_be_clickable((elementIndex, value)))
            self.driver.switch_to.frame(iframe)
        except:
            print(f'Error:To switch To Frame !!! ({by}:{value})')

    def elementText(self,by,value):
        elementIndex = self.__byIndexSelect(by=by)
        try: 
            element = self.wait.until(EC.element_to_be_clickable((elementIndex, value)))
            return element.text
        except:
            print(f'Error:To find elemnet !!! ({by}:{value})')
            return
        
    def elementValue(self,by,value):
        elementIndex = self.__byIndexSelect(by=by)
        try: 
            element = self.wait.until(EC.element_to_be_clickable((elementIndex, value)))
            return element.get_attribute('value')
        except:
            print(f'Error:To find elemnet !!! ({by}:{value})')
            return

    def tearDown(self):
        self.driver.quit()
        


# html body.has-product-menu-bar
# webControl.tearDown() 

