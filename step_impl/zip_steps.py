from getgauge.python import step
import subprocess
import uiautomation

@step("Open 7zip")
def open_7zip():
    subprocess.Popen("C:\\Program Files\\7-Zip\\7zFM.exe")

@step("Close 7zip")
def close_7zip():
    zipWindow = uiautomation.WindowControl(Name="7-Zip")

    uiautomation.WaitForExist(zipWindow, 5)
    zipWindow.MenuItemControl(Name="File").Click()
    
    # Because Application changes after clicking file
    zipWindow = uiautomation.WindowControl(Name="7-Zip")
    uiautomation.WaitForExist(zipWindow, 5)
    zipWindow.MenuItemControl(RegexName="Exit.*").Click()

@step("Go to <menu> <submenu> in 7-zip")
def go_to(menuItem, submenuItem):
    zipWindow = uiautomation.WindowControl(Name="7-Zip")

    uiautomation.WaitForExist(zipWindow, 5)
    zipWindow.MenuItemControl(RegexName=f"{menuItem}.*").Click()
    
    # Because Application changes after clicking file
    zipWindow = uiautomation.WindowControl(Name="7-Zip")
    uiautomation.WaitForExist(zipWindow, 5)
    zipWindow.MenuItemControl(RegexName=f"{submenuItem}.*").Click()

@step("Click <button> button in 7-zip")
def click_button(button):
    zipWindow = uiautomation.WindowControl(Name="7-Zip")

    uiautomation.WaitForExist(zipWindow, 5)
    zipWindow.ButtonControl(Name=button).Click()



@step("Verify About 7-zip contains <7-Zip is free software> text")
def etozvx(text):
    zipWindow = uiautomation.WindowControl(Name="7-Zip")

    uiautomation.WaitForExist(zipWindow, 5)
    assert zipWindow.TextControl(RegexName=f"{text}.*").Exists(), f"Text '{text}' does not exist in the about window or About is not opened"


