set ie= createobject("InternetExplorer.Application")
ie.visible=true
ie.navigate "https://opensource-demo.orangehrmlive.com/"
ie.FullScreen=true
ie.AddressBar = True  
ie.StatusBar = False
ie.MenuBar = True

Do until ie.readystate=4
Loop


set oUser = ie.Document.GetElementById("txtUsername")
oUser.value="Admin"

set oPassword= ie.Document.GetElementById("txtPassword")
oPassword.value="admin123"

set oLoginButton= ie.Document.getElementById("btnLogin")
oLoginButton.click

WScript.sleep 5000
Do until ie.readystate=4
Loop
set oMyInfo = ie.Document.getElementById("menu_pim_viewPimModule")
oMyInfo.click
