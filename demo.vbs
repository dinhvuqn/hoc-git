Dim objHTTP
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

' Specify the URL to request
url = "https://www.selenium.dev/documentation/webdriver/"

' Open a connection to the specified URL
objHTTP.Open "GET", url, False

' Send the HTTP request
objHTTP.Send

' Display the response text
MsgBox objHTTP.responseText

' Clean up
Set objHTTP = Nothing
