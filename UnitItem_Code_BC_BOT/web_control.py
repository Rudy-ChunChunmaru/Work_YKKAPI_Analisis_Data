from selenium import webdriver

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC



class WebControlBC:
    def __init__(self) -> None:
        self.driver = webdriver.Firefox()


test = WebControlBC
test.driver.get('https://businesscentral.dynamics.com')

# elementsignin = WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.ID, "c-shellmenu_custom_button_outline_newtab_signin_bhvr100_right")))
# elementsignin.click()

elementemail = driver.find_element(By.NAME,"loginfmt")
elementemail.clear()
elementemail.send_keys("pycon")
elementemail.send_keys(Keys.RETURN)

driver.close()