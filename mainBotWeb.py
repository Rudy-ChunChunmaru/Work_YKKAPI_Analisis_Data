from bot.web_control import WebControlBC

user = 'admin@ykkapindonesia.onmicrosoft.com'
password = '88{aW<5%)T'

def startWebConrtol():
    return WebControlBC('https://businesscentral.dynamics.com/YKK')

def loginToBC(webControl):
    webControl.clickWeb(by='id',value='c-shellmenu_custom_button_outline_newtab_signin_bhvr100_right')
    webControl.switchToSecondWeb()
    webControl.inputWeb(by='name',value='loginfmt',inputValue=user)
    webControl.clickWeb(by='id',value='idSIButton9')
    webControl.inputWeb(by='name',value='passwd',inputValue=password)
    webControl.clickWeb(by='id',value='idSIButton9')
    webControl.clickWeb(by='id',value='idSIButton9')
    webControl.clickWeb(by='classSelector',value='#apps-module-banner > div:nth-child(2) > button:nth-child(1)')
    webControl.clickWeb(by='classSelector',value='div.___11eoy74:nth-child(3)')
    webControl.switchToSecondWeb()
    webControl.clickWeb(by='classSelector',value='.product-name')
    webControl.clickWeb(by='classSelector',value='.environmentPickerButton--lPstm2EbOaZnwaOHZEcU')

def setupEnvBC(webControl,adddressEnv):
    webControl.switchToSecondWeb()
    webControl.clickWeb(by='classSelector',value='.product-name')
    webControl.clickWeb(by='classSelector',value='.ms-Button')
    webControl.redirectWeb(value=adddressEnv)


# TODO-1: Start Web Control
webControl = startWebConrtol()
loginToBC(webControl)

# TODO-2: set Env to uat
EnvBC = 'https://businesscentral.dynamics.com/uat?company=PT%20YKK%20AP'
setupEnvBC(webControl,adddressEnv=EnvBC)

# TODO-3: go to AO 
NodeAO = '&node=0000e6dc-40c8-0000-1013-0300836bd2d2'
webControl.redirectWeb(value=EnvBC+NodeAO)


