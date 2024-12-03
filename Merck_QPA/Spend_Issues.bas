Attribute VB_Name = "Spend_Issues"
Sub SpendIssues()

Dim Driver As WebDriver


Sheets(1).Activate
LastCount = Range("B1").End(xlDown).Row

'Starting Chrome Web Driver
Set Driver = CreateObject("Selenium.ChromeDriver")
Driver.Start "chrome"
    
'Maximizing Chrome Window and Opening URL
Driver.Window.Maximize

Driver.Get "https://login.medispend.com/dashboard/login"
    
Application.Wait Now + TimeValue("00:00:05")
    
'Entering Credentails - Email Only
Driver.FindElementById("1-email").SendKeys Sheets("Sheet1").Range("B1").Value
Driver.FindElementByClass("auth0-label-submit").Click

WaitForElementByXPath Driver, "//button[normalize-space()='Change My Site']"
Application.Wait Now + TimeValue("00:00:01")

Call GenerateSpendIssue(Driver)

Call GenerateCrIssue(Driver)

Driver.Quit
MsgBox "Completed", vbInformation, "Success"

End Sub
Private Sub GenerateSpendIssue(Driver As WebDriver)

'Navigating to Dashboard
Driver.FindElementByXPath("//a[normalize-space()='Dashboards']").Click
Application.Wait Now + TimeValue("00:00:01")
Driver.FindElementByXPath("//a[normalize-space()='Resolution Center']").Click

'Waiting for Loading box to disapper
WaitForElementByXPath Driver, "//input[@id='myWorkTab_assignedItemsDueFrom']"

Driver.FindElementByXPath("//a[@href='#spendIssuesTab']").Click
WaitForElementByXPath Driver, "//*[@id='myDropdown']/button[1]"
Driver.FindElementByXPath("//button[@class='btn btn-default btn-xs buttonExportBulk']").Click

'Waiting for Processing
Status = Driver.FindElementByXPath("*//table/tbody/tr[1]/td[5]").Text
While Status = "Processing..."
   Application.Wait Now + TimeValue("00:00:02")
   Status = Driver.FindElementByXPath("*//table/tbody/tr[1]/td[5]").Text
Wend

'Clicking on Download
Application.Wait Now + TimeValue("00:00:04")
WaitForElementByXPath Driver, "//*//table/tbody/tr[1]/td[4]/a[1]"
Driver.FindElementByXPath("*//table/tbody/tr[1]/td[4]/a[1]").Click
Application.Wait Now + TimeValue("00:00:10")

downloadFolderPath = Environ("USERPROFILE") & "\Downloads\"
                
Do While Dir(downloadFolderPath & "*.crdownload") <> ""
    Application.Wait Now + TimeValue("00:00:01")
Loop

End Sub
Private Sub GenerateCrIssue(Driver As WebDriver)

Driver.FindElementByXPath("//a[normalize-space()='Dashboards']").Click
Application.Wait Now + TimeValue("00:00:01")
Driver.FindElementByXPath("//a[normalize-space()='Resolution Center']").Click

'Waiting for Loading box to disapper
WaitForElementByXPath Driver, "//input[@id='myWorkTab_assignedItemsDueFrom']"

Driver.FindElementByXPath("//a[@href='#crIssuesTab']").Click
WaitForElementByXPath Driver, "//*[@id='myDropdown']/button[1]"
Driver.FindElementByXPath("//button[@id='exportCRIssues']").Click

'Waiting for Processing
Status = Driver.FindElementByXPath("*//table/tbody/tr[1]/td[5]").Text
While Status = "Processing..."
   Application.Wait Now + TimeValue("00:00:02")
   Status = Driver.FindElementByXPath("*//table/tbody/tr[1]/td[5]").Text
Wend

'Clicking on Download
Application.Wait Now + TimeValue("00:00:04")
WaitForElementByXPath Driver, "//*//table/tbody/tr[1]/td[4]/a[1]"
Driver.FindElementByXPath("*//table/tbody/tr[1]/td[4]/a[1]").Click
Application.Wait Now + TimeValue("00:00:10")

downloadFolderPath = Environ("USERPROFILE") & "\Downloads\"
                
Do While Dir(downloadFolderPath & "*.crdownload") <> ""
    Application.Wait Now + TimeValue("00:00:01")
Loop

End Sub
Private Function WaitForElementByXPath(Driver As WebDriver, XPath As String)

Dim Element As WebElement
While Driver.FindElementsByXPath(XPath).Count = 0
    Application.Wait Now + TimeValue("00:00:01")
Wend

Set Element = Driver.FindElementByXPath(XPath)

For i = 1 To 30
    If Element.IsDisplayed And Element.IsEnabled Then
        Exit For
    End If
    Application.Wait Now + TimeValue("00:00:01")
    Set Element = Driver.FindElementByXPath(XPath)
Next i

End Function
