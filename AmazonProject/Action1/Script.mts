SystemUtil.CloseProcessByName"Chrome.exe"
'SystemUtil.Run"Chrome.exe","www.amazon.in"

'Function TC_02()
'
'Dim searchString,res1
'searchString="Inverters"
'SignIn()
'Browser(browserObject).Page(pageObject).WebList("html id:=searchDropdownBox").Select "Home & Kitchen"
'Browser(browserObject).Page(pageObject).WebButton("name:=Go").Click
'Browser(browserObject).Page(pageObject).Link(HAndkAppliances).Click
'Browser(browserObject).Page(pageObject).WebElement("innerhtml:=Inverters").Click
'wait(5)
'res1=Browser(browserObject).Page(pageObject).WebElement("xpath:=//LI/SPAN[normalize-space()='Inverters']/SPAN[1]").GetROProperty("text")
'
'If inStr(searchString,res1)>=0 Then
'	Reporter.ReportEvent miccPass,"verifyInverters","Pass"
'Else
'	Reporter.ReportEvent miccPass,"verifyInverters","Fail"
'End If
'SignOut()
'SystemUtil.CloseProcessByName"Chrome.exe"
'End Function



DataTable.AddSheet "Test Data"
DataTable.ImportSheet "C:\Users\user240\Documents\Amazon\Test_Data\Test Data.xlsx","Amazon Data","Test Data"
rowCount = DataTable.GetSheet("Test Data").GetRowCount

For i= 1 To rowCount
DataTable.SetCurrentRow (i)
If DataTable.Value("Expected_Value","Test Data")="Y" Then
  SystemUtil.Run"Chrome.exe","www.amazon.in"

 ExecuteTest(DataTable.Value("testCaseID","Test Data"))
'  ExecuteTest "TC_001"

'Environment.Value("Result")="Pass"
DataTable.Value("Result","Test Data")=Environment.Value("Result")
End If
Next
DataTable.ExportSheet "C:\Users\user240\Documents\Amazon\Test_Data\Test Data.xlsx","Test Data","Amazon Data"


'SignIn()

'SignOut() @@ script infofile_;_ZIP::ssf98.xml_;_

'TC_01()

'TC_02()

'TC_03()

'TC_04()

'TC_05()

'TC_06()
 @@ script infofile_;_ZIP::ssf171.xml_;_
' TC_07()

' TC_08()
 
' TC_09()
 
  'TC_10()



