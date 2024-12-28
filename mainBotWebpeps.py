from bot.web_control import WebControlBC
import time

user = ''
password = ''

def startWebConrtol():
    return WebControlBC('https://businesscentral.dynamics.com/YKK')

def loginToBC(webControl):
    webControl.clickWeb(by='id',value='c-shellmenu_custom_button_outline_newtab_signin_bhvr100_right')
    webControl.switchToSecondWeb()
    time.sleep(0.5)
    webControl.inputWeb(by='name',value='loginfmt',inputValue=user)
    webControl.clickWeb(by='id',value='idSIButton9')
    time.sleep(0.5)
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

def openAO(EnvBC,webControl):
    Node = '&node=0000e6dc-40c8-0000-1013-0300836bd2d2'
    webControl.redirectWeb(value=EnvBC+Node)

def openRecevingAO(EnvBC,webControl):
    Node='&node=0000f80d-c80b-0000-102b-3500836bd2d2'
    webControl.redirectWeb(value=EnvBC+Node)

def selectAONumber(webControl,number):
    webControl.switchToFrame(by='classSelector',value='.designer-client-frame')
    webControl.clickWeb(by='classSelector',value='.search-box-container--GSU2NCtWS0dJ4gDmDdR4')
    webControl.clickWeb(by='classSelector',value='.search-box-container--GSU2NCtWS0dJ4gDmDdR4')
    webControl.clickWeb(by='classSelector',value='.fui-Input__input')

    while webControl.elementValue(by='classSelector',value='.fui-Input__input') != number:
        webControl.inputWeb(by='classSelector',value='.fui-Input__input',inputValue=number)
        time.sleep(0.25)
    while webControl.elementText(by='classSelector',value='#b1tee') != number:
        time.sleep(0.25)
    webControl.clickWeb(by='classSelector',value='#b6aee')

# TODO-1: Start Web Control
webControl = startWebConrtol()
loginToBC(webControl)

# TODO-2: set Env to uat
EnvBC = 'https://businesscentral.dynamics.com/uat?company=PT%20YKK%20AP'
setupEnvBC(webControl,adddressEnv=EnvBC)


while True:
    print('''
    Menu:
    1. AO
    2. Receving AO
    3. List Breakdown AO
    4. Stop
    ''')
    Menu = input('menu type ?\n')
    if(Menu == '4'):
        webControl.tearDown()
        break
    else:
        Number_AO = input('next AO to Find ?\n')
        if(Number_AO == '0'):
            webControl.tearDown()
            break
        else:
            if Menu == '1':
                # TODO-3: go to AO 
                openAO(EnvBC,webControl)
                # TODO-4: select AO number
                selectAONumber(webControl,Number_AO)
                continue
            elif Menu == '2':
                # TODO-3: go to Receving AO 
                openRecevingAO(EnvBC,webControl)
                # TODO-4: select AO number
                selectAONumber(webControl,Number_AO)
                continue





