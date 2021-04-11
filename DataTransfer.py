from tkinter import *
import tkinter as tk
from tkinter import messagebox
import time, configparser, re, webbrowser, pyautogui


#Define function to read config file
def readConfig():
    config = configparser.ConfigParser()
    config.read('config.ini')
    return config

# Refer to config file to change the duration
WAIT_INFO = int(readConfig()['SECONDS']['waitforinfotab'])
WAIT_TEST_PAGE = int(readConfig()['SECONDS']['waitfortestpage'])
WAIT_TEST_TYPE = int(readConfig()['SECONDS']['waitfortesttype'])

#Refer to config file to change the click position
#temperatureX=int(readConfig()['CLICKS']['temperatureX'])
#temperatureY=int(readConfig()['CLICKS']['temperatureY'])
#startX=int(readConfig()['CLICKS']['startX'])
#startY=int(readConfig()['CLICKS']['startY'])
#uutX=int(readConfig()['CLICKS']['uutX'])
#uutY=int(readConfig()['CLICKS']['uutY'])


#Dropdown List for Amendment
f=open("Amendment.txt", "r")
AMENDMENT=[]
for line in f:
    line=line.rstrip()
    AMENDMENT.append(line)
f.close()


#Dropdown List for Operator Name
f=open("NameList.txt", "r")
OPERATOR=[]
for line in f:
    line=line.rstrip()
    OPERATOR.append(line)
f.close()

#Dropdown List for Fixture Name
f=open("Fixture.txt", "r")
FIXTURE=[]
for line in f:
    line=line.rstrip()
    FIXTURE.append(line)
f.close()

#Dropdown List for Tua Name
f=open("Tua.txt", "r")
TUA=[]
for line in f:
    line=line.rstrip()
    TUA.append(line)
f.close()

#Dropdown List for Temperature
TEMPERATURE = ["25 C","-40 C","85 C"]

#Dropdown List for Product Name
f=open("Product.txt","r")
PRODUCT=[]
for line in f:
    line=line.rstrip()
    PRODUCT.append(line)
f.close()

#Dropdown List for Product Name
#PRODUCT=['E36170DA','E36170KA','E36167DA01','E36167GA01','E36193DA01']

#Dropdown List for Test Type
TEST=['Reinforced','Normal']


root= tk.Tk()
root.geometry('400x500')
root.title('DataTransfer ' + readConfig()['PROGRAM']['version'])
root.lift()
root.attributes('-topmost',True)



def open_readme():
        webbrowser.open('readme.txt')


#To ensure latest info is displayed
def Productcallback(*root):
    Product.get()

#To ensure latest info is displayed
def Amendmentcallback(*root):
    Amendment.get()

#To ensure latest info is displayed
def Operatorcallback(*root):
    Operator.get()

#To ensure latest info is displayed
def Fixturecallback(*root):
    Fixture.get()

#To ensure latest info is displayed
def Tuacallback(*root):
    Tua.get()
    
#To ensure latest info is displayed
def Temperaturecallback(*root):
    Temperature.get()

#To select type of test
def Testcallback(*root):
    Test.get()


def ClickStart():
    try:
        pyautogui.moveTo(389,46,duration=0.1)
        pyautogui.click()
    except:
        pass

def ClickTemperaturePageSBE1():
    try:
        pyautogui.moveTo(762,368,duration=0.1)
        pyautogui.click()
    except:
        pass

def ClickTemperaturePageSBE2():
    try:
        pyautogui.moveTo(889,404,duration=0.1)
        pyautogui.click()
    except:
        pass

def ClickInfo():
    try:
        pyautogui.moveTo(731,387,duration=0.1)
        pyautogui.click()
    except:
        pass
    
#Select Posage 2 Normal
def SelectPosage2Normal():
    pyautogui.hotkey('tab','tab')
    pyautogui.press('down')
    pyautogui.hotkey('enter')
    
#Select Posage 2 Reinforced
def SelectPosage2Reinforced():
    pyautogui.hotkey('tab','tab')
    pyautogui.press('down','down')
    pyautogui.press('down')
    pyautogui.hotkey('enter')


#To clear S/N entry box after all data has been transferred
def ClearText():
    SerialNumber.delete(0,END)

#Enter all SBE1 info into the infobox
def EnterSBE1Info():
    pyautogui.moveTo(1032,312,duration=0.1)
    pyautogui.click()
    time.sleep(2)
    checkAmendment()
    time.sleep(2)
    pyautogui.moveTo(1032,341,duration=0.1)
    pyautogui.click()
    time.sleep(2)
    # To erase unwanted detals in S/N section
    pattern = r'[HE]'
    mod_string = re.sub(pattern, '', SerialNumber.get())
    pyautogui.typewrite(mod_string)
    time.sleep(2)
    pyautogui.moveTo(1032,410,duration=0.1)
    pyautogui.click()
    time.sleep(2)
    pyautogui.moveTo(1032,410,duration=0.1)
    pyautogui.click()
    checkOperator()
    time.sleep(2)
    pyautogui.moveTo(1032,446,duration=0.1)
    pyautogui.click()
    time.sleep(2)
    checkTemperature()
    time.sleep(2)
    pyautogui.moveTo(720,510,duration=0.1)
    pyautogui.click()

#Enter all SBE2 info into the infobox
def EnterSBE2Info():
    pyautogui.moveTo(1032,312,duration=0.1)
    pyautogui.click()
    time.sleep(2)
    checkAmendment()
    time.sleep(2)
    pyautogui.moveTo(1032,341,duration=0.1)
    pyautogui.click()
    time.sleep(2)
    # To erase unwanted detals in S/N section
    pattern = r'[HE]'
    mod_string = re.sub(pattern, '', SerialNumber.get())
    pyautogui.typewrite(mod_string)
    time.sleep(2)
    pyautogui.moveTo(1032,410,duration=0.1)
    pyautogui.click()
    time.sleep(2)
    pyautogui.moveTo(1032,410,duration=0.1)
    pyautogui.click()
    checkOperator()
    time.sleep(2)
    pyautogui.moveTo(1032,446,duration=0.1)
    pyautogui.click()
    time.sleep(2)
    checkTemperature()
    time.sleep(2)
    pyautogui.moveTo(720,510,duration=0.1)
    pyautogui.click()

#Enter all EDU info into the infobox
def EnterEDUInfo():
    pyautogui.moveTo(1032,312,duration=0.1)
    pyautogui.click()
    time.sleep(2)
    checkAmendment()
    time.sleep(2)
    pyautogui.moveTo(1032,341,duration=0.1)
    pyautogui.click()
    time.sleep(2)
    # To erase unwanted detals in S/N section
    pattern = SerialNumber.get()
    mod_string = pattern.replace('E361930', '' )
    pyautogui.typewrite(mod_string,interval=0.5)
    time.sleep(2)
    pyautogui.moveTo(1032,410,duration=0.1)
    pyautogui.click()
    time.sleep(2)
    pyautogui.moveTo(1032,410,duration=0.1)
    pyautogui.click()
    checkOperator()
    time.sleep(2)
    pyautogui.moveTo(1032,446,duration=0.1)
    pyautogui.click()
    time.sleep(2)
    checkTemperature()
    time.sleep(2)
    pyautogui.moveTo(720,510,duration=0.1)
    pyautogui.click()

        
my_menu = Menu(root)
root.config(menu=my_menu)


#Create a menu item
file_menu = Menu(my_menu)
my_menu.add_cascade(label='File', menu=file_menu)
file_menu.add_separator()
file_menu.add_command(label='Exit', command=root.quit)

# Create a help menu 
help_menu= Menu(my_menu)
my_menu.add_cascade(label='Help', menu=help_menu)
help_menu.add_command(label='About', command=open_readme)

#Frame 'Amendment'
mylabel1 = Label(root, text="Amendment")
mylabel1.grid(row=0, column=0)

#Entry : Amendment
Amendment=StringVar(root)
menu=OptionMenu(root,Amendment, *AMENDMENT, command=Amendmentcallback)
menu.grid(row=2, column=0)
Amendment.set('.')

#Frame ' Serial Number'
mylabel2 = Label (root, text="Serial Number")
mylabel2.grid(row=4, column=0)

#Entry: Serial Number
SerialNumber =  Entry(root)
SerialNumber.focus()
SerialNumber.grid(row=5, column=0)

#Frame ' Operator '
mylabel3 = Label (root, text="Operator Name")
mylabel3.grid(row=7, column=0)

#Entry: Operator Name
Operator=StringVar(root)
menu=OptionMenu(root,Operator, *OPERATOR, command=Operatorcallback)
menu.grid(row=8, column=0)
Operator.set('NA')

#Frame ' Fixture'
mylabel3 = Label (root, text="Fixture")
mylabel3.grid(row=7, column=2)

#Entry: Fixture Number
Fixture=StringVar(root)
menu=OptionMenu(root,Fixture, *FIXTURE, command=Fixturecallback)
menu.grid(row=8, column=2)
Fixture.set('NA')

#Frame ' Tua'
mylabel3 = Label (root, text="Tua")
mylabel3.grid(row=7, column=4)

#Entry: Tua Number
Tua=StringVar(root)
menu=OptionMenu(root,Tua, *TUA, command=Tuacallback)
menu.grid(row=8, column=4)
Tua.set('NA')

#Frame ' Temperature'
mylabel4 = Label (root, text="Temperature")
mylabel4.grid(row=10, column=0)

#Entry: Temperature
Temperature=StringVar(root)
menu=OptionMenu(root,Temperature, *TEMPERATURE, command=Temperaturecallback)
menu.grid(row=11, column=0)
Temperature.set('25 C')

#Frame ' Product '
mylabel5 = Label (root, text="Part Number")
mylabel5.grid(row=13, column=0)

#Entry: 'Product Name'
Product=StringVar(root)
menu=OptionMenu(root,Product, *PRODUCT, command=Productcallback)
menu.grid(row=14, column=0)
Product.set('E36170DA')

#Frame 'Test Type'
mylabel6 = Label(root, text='Test Type(SBE2 ONLY)')
mylabel6.grid(row=16, column=0)

#Entry: Renforced or Normal
Test=StringVar(root)
menu=OptionMenu(root,Test, *TEST, command=Testcallback)
menu.grid(row=17, column=0)
Test.set('Normal')

def checkAmendment():
    pyautogui.hotkey('enter')
    pyautogui.typewrite(Amendment.get())

def checkOperator():
    pyautogui.hotkey('enter')
    pyautogui.typewrite(Operator.get()+'-'+Fixture.get() +'_' + Tua.get())

def checkTemperature():
    pyautogui.hotkey('enter')
    pyautogui.typewrite(Temperature.get())

def checkSN():
      value=Product.get()
      if( value=='E36170DA' )or (value=='E36170KA'):
          SerialNumberRe = re.compile(r'^\w{6}$')
          SerialNumberMO = SerialNumberRe.search(SerialNumber.get())
          if Temperature.get()=='25 C':
             ClickTemperaturePageSBE1()
             time.sleep(1)
             #Tap once to select 25C
             pyautogui.hotkey('tab','tab')
             time.sleep(1)
             pyautogui.hotkey('enter')
             #delay to wait for Info Page to appear
             time.sleep(WAIT_INFO)
             EnterSBE1Info()
             #delay added to wait for Test Initialisation page to appear
             time.sleep(WAIT_TEST_PAGE)
             ClickStart()
             ClearText()
             
          if Temperature.get()=='-40 C':
              ClickTemperaturePageSBE1()
              time.sleep(1)
              #Double tap to select desired Temperature
              pyautogui.hotkey('tab','tab')
              time.sleep(1)
              pyautogui.press('down')
              time.sleep(1)
              pyautogui.press('down')
              time.sleep(1)
              pyautogui.hotkey('enter')
              #Added delay to wait for Info page to appear
              time.sleep(WAIT_INFO)
              EnterSBE1Info()
              # delay added to wait for Test Initialisation page to appear
              time.sleep(WAIT_TEST_PAGE)
              ClickStart()
              ClearText()
             
          if Temperature.get()=='85 C':
              ClickTemperaturePageSBE1()
              time.sleep(1)
              # Double tap to select desired Temperature
              pyautogui.hotkey('tab','tab')
              time.sleep(1)
              pyautogui.press('down')
              time.sleep(1)
              pyautogui.hotkey('enter')
              # Added delay to wait for Info page to appear
              time.sleep(WAIT_INFO)
              EnterSBE1Info()
              # delay added to wait for Test Initialisation page to appear
              time.sleep(WAIT_TEST_PAGE)
              ClickStart()
              ClearText()
             
           
            
            

      elif( value=='E36167DA01') or (value=='E36167GA01'):
          SerialNumberRe = re.compile(r'^\w{5}$')
          SerialNumberMO = SerialNumberRe.search(SerialNumber.get())
          if Temperature.get()=='25 C':
              ClickTemperaturePageSBE2()
              time.sleep(1)
              #Select 25C
              pyautogui.hotkey('tab','tab')
              time.sleep(1)
              pyautogui.press('down')
              time.sleep(1)
              pyautogui.press('down')
              time.sleep(1)
              pyautogui.hotkey('enter')
              # Added delay to wait for test type selection
              time.sleep(WAIT_TEST_TYPE)
              if Test.get()=='Normal':
                    SelectPosage2Normal()
                # Added delay to wait for Info page to appear
                    time.sleep(WAIT_INFO)
                    EnterSBE2Info()
                # delay added to wait for Test Initialisation page to appear
                    time.sleep(WAIT_TEST_PAGE)
                    ClickStart()
                    ClearText()
              else:
                    SelectPosage2Reinforced()
                # Added delay to wait for Info page to appear
                    time.sleep(WAIT_INFO)
                    EnterSBE2Info()
                # delay added to wait for Test Initialisation page to appear
                    time.sleep(WAIT_TEST_PAGE)
                    ClickStart()
                    ClearText()
             
          if Temperature.get()=='-40 C':
                # Select -40C
                ClickTemperaturePageSBE2()
                time.sleep(1)
                pyautogui.hotkey('tab','tab')
                time.sleep(1)
                pyautogui.press('down')
                time.sleep(1)
                pyautogui.hotkey('enter')
                # Added delay to wait for test type selection
                time.sleep(WAIT_TEST_TYPE)
                if Test.get()=='Normal':
                                SelectPosage2Normal()
                # Added delay to wait for Info page to appear
                                time.sleep(WAIT_INFO)
                                EnterSBE2Info()
                # delay added to wait for Test Initialisation page to appear
                                time.sleep(WAIT_TEST_PAGE)
                                ClickStart()
                                ClearText()
                else:
                                    SelectPosage2Reinforced()
                # Added delay to wait for Info page to appear
                                    time.sleep(WAIT_INFO)
                                    EnterSBE2Info()
                # delay added to wait for Test Initialisation page to appear
                                    time.sleep(WAIT_TEST_PAGE)
                                    ClickStart()
                                    ClearText()


             
          if Temperature.get()=='85 C':
                # Select 85C
                ClickTemperaturePageSBE2()
                time.sleep(1)
                pyautogui.hotkey('tab','tab')
                time.sleep(1)
                pyautogui.hotkey('enter')
                # Added delay to wait for test type selection
                time.sleep(WAIT_TEST_TYPE)
                if Test.get()=='Normal':
                                SelectPosage2Normal()
                                # Added delay to wait for Info page to appear
                                time.sleep(WAIT_INFO)
                                EnterSBE2Info()
                                # delay added to wait for Test Initialisation page to appear
                                time.sleep(WAIT_TEST_PAGE)
                                ClickStart()
                                ClearText()
                else:
                                SelectPosage2Reinforced()
                                # Added delay to wait for Info page to appear
                                time.sleep(WAIT_INFO)
                                EnterSBE2Info()
                                # delay added to wait for Test Initialisation page to appear
                                time.sleep(WAIT_TEST_PAGE)
                                ClickStart()
                                ClearText()

            

            
      elif( value=='E36193DA01' ):
           SerialNumberRe = re.compile(r'^\w{10}$')
           SerialNumberMO = SerialNumberRe.search(SerialNumber.get())       
           if Temperature.get()=='25 C':
             
               EnterEDUInfo()
               #delay added to wait for temperature selection page to load
               time.sleep(WAIT_INFO)
               pyautogui.hotkey('tab','tab')
               time.sleep(1)
               pyautogui.press('down')
               time.sleep(1)
               pyautogui.press('down')
               time.sleep(1)
               pyautogui.hotkey('enter')
               # delay added to wait for Test Initialisation page to appear
               time.sleep(WAIT_TEST_PAGE)
               ClickStart()
               ClearText()
             
           if Temperature.get()=='-40 C':
                 
                 EnterEDUInfo()
             # delay added to wait for temperature selection page to load
                 time.sleep(WAIT_INFO)
                 pyautogui.hotkey('tab','tab')
                 time.sleep(1)
                 pyautogui.press('down')
                 time.sleep(1)
                 pyautogui.hotkey('enter')
             # delay added to wait for Test Initialisation page to appear
                 time.sleep(WAIT_TEST_PAGE)
                 ClickStart()
                 ClearText()

           if Temperature.get()=='85 C':
                 
                 EnterEDUInfo()
             # delay added to wait for temperature selection page to load
                 time.sleep(WAIT_INFO)
                 pyautogui.hotkey('tab','tab')
                 time.sleep(1)
                 pyautogui.hotkey('enter')
             # delay added to wait for Test Initialisation page to appear
                 time.sleep(WAIT_TEST_PAGE)
                 ClickStart()
                 ClearText()
      else:
        messagebox.showwarning('SN Entry',
                                   'Invalid Serial number. Try again.')
          
      return True



def startApp():
    checkSN()
        

# Button : 'Start
start_btn = tk.Button(root, text='Start', command=startApp)
start_btn.grid(columnspan=4, row=19, column=1, pady=10)


root.mainloop()

