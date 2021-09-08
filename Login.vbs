DIM oIE
DIM ipf

Set oIE = CreateObject("InternetExplorer.Application")
oIE.navigate "http://www.gmail.com"
oIE.Visible = True

While oIE.Busy
     WScript.Sleep 50
Wend

Set UID = oIE.document.all.email
UID.value = "narurathore"

Set PWD = oIE.document.all.Passwd
PWD.value = "johndalton"

oIE.document.all.Signin.click