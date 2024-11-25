from selenium import webdriver

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class WebControlBC:
    def __init__(self,adderess,valueWait=10) -> None:
        self.driver = webdriver.Firefox()
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
        except:
            print(f'Error:To find elemnet !!! ({by}:{value})')


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
    
    def tearDown(self):
        self.driver.quit()
        


# html body.has-product-menu-bar
# webControl.tearDown() 

