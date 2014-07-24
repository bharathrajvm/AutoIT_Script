
;*******************************************************************
;Description: SSL Overloading
;
;Purpose:
;
;Date: 18 Jun 2014 , Modified on 19 June 2014
;Author: Ganesh
;Company: Brillio
;*********************************************************************

;********************************************
;Include Standard Library

;*******************************************
#include <Constants.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <IE.au3>
#include <Clipboard.au3>
#include <Date.au3>
#include <GuiListView.au3>
#include <GUIConstantsEx.au3>
#include <GuiTreeView.au3>
#include <GuiImageList.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <GuiTreeView.au3>
#include <File.au3>
;******************************************

;***************************************************************
;Initialize AutoIT Key delay
;****************************************************************
AutoItSetOption ( "SendKeyDelay", 400)

;******************************************************************
;Reading test data from xls
;To do - move helper function
;******************************************************************
;Open xls
Local $sFilePath1 =  @ScriptDir & "\" & "TestData.xlsx";This file should already exist in the mentioned path
Local $oExcel = _ExcelBookOpen($sFilePath1,0,True)

;Check for error
If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "File does not exist")
    Exit
 EndIf

 ; Reading xls data into variables
;to do - looping to get the data from desired row of xls
;to do - looping to get the data from desired row of xls
Local $testCaseIteration = _ExcelReadCell($oExcel, 16, 1)
Local $testCaseExecute = _ExcelReadCell($oExcel, 16, 2)
Local $testCaseName = _ExcelReadCell($oExcel, 16, 3)
Local $testCaseDescription = _ExcelReadCell($oExcel, 16, 4)
Local $JunoOrKep  = _ExcelReadCell($oExcel, 16, 5)
Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 16, 6)
;if $JunoOrKep = "Juno" Then
  ; Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 16, 6)
;Else
   ;Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 16, 7)
   ;EndIf
Local $testCaseWorkSpacePath = _ExcelReadCell($oExcel, 16, 8)
Local $testCaseProjectName = _ExcelReadCell($oExcel, 16, 9)
Local $testCaseJspName = _ExcelReadCell($oExcel, 16, 10)
Local $testCaseJspText = _ExcelReadCell($oExcel, 16, 11)
Local $testCaseAzureProjectName = _ExcelReadCell($oExcel, 16, 12)
Local $testCaseCheckJdk = _ExcelReadCell($oExcel, 16, 13)
Local $testCaseJdkPath = _ExcelReadCell($oExcel, 16, 14)
Local $testCaseCheckLocalServer = _ExcelReadCell($oExcel, 16, 15)
Local $testCaseServerPath = _ExcelReadCell($oExcel, 16, 16)
Local $testCaseServerNo = _ExcelReadCell($oExcel, 16, 17)
Local $testCaseUrl = _ExcelReadCell($oExcel, 16, 19)
Local $testCaseValidationText = _ExcelReadCell($oExcel, 16, 19)
Local $testCaseSubscription = _ExcelReadCell($oExcel, 16, 20)
Local $testCaseStorageAccount = _ExcelReadCell($oExcel, 16, 21)
Local $testCaseServiceName = _ExcelReadCell($oExcel, 16, 22)
Local $testCaseTargetOS = _ExcelReadCell($oExcel, 16, 23)
Local $testCaseTargetEnvironment = _ExcelReadCell($oExcel, 16, 24)
Local $testCaseCheckOverwrite = _ExcelReadCell($oExcel, 16, 25)
Local $testCaseJDKOnCloud = _ExcelReadCell($oExcel, 16, 28)
Local $testCaseUserName = _ExcelReadCell($oExcel, 16, 29)
Local $testCasePassword = _ExcelReadCell($oExcel, 16, 30)
Local $testcaseNewSessionJSPText = _ExcelReadCell($oExcel, 16, 31)
Local $testcaseExternalJarPath = _ExcelReadCell($oExcel, 16, 32)
Local $testcaseCertificatePath = _ExcelReadCell($oExcel, 16, 33)
_ExcelBookClose($oExcel,0)
;*******************************************************************************

;to do - Pre validation steps

;Opening instance of Eclipse
OpenEclipse()

Delete()
;Creating Java Project
CreateJavaProject()

;Creating JSP file and insert code
CreateJSPFile()

;Create Azure Package
CreateAzurePackage()

;Enable SSL Offloading
EnableSSLOffloading()

;Publish to Cloud
PublishToCloud()

;Wait for 10 min RDP screen
Sleep(540000)

;Check RDP and Open excel
CheckRDPConnection()

;Check for published key word in Azure activity log and update excel
ValidateTextAndUpdateExcel()

;to do - Post validation steps


If ProcessExists("eclipse.exe") Then
		  ProcessClose("eclipse.exe")
EndIf

;***************************************************************
;Helper Functions
;***************************************************************

Func wincheck($fun ,$ctrl)
 	  Local $act = WinActive($ctrl)
	  if $act = 0 Then
		 Local $lFile = FileOpen(@ScriptDir & "\" & "Error.log", 1)
 		 Local $wrt = _FileWriteLog(@ScriptDir & "\" & "Error.log", "Error Opening:" & $ctrl, 1)
		 MsgBox("","",$wrt)
		 FileClose($lFile)
		 MsgBox($MB_OK,"Error","Error status is recorded in Error.log")
	  EndIf
 EndFunc

;***************************************************************
;Function to Open instance of Eclipse
;***************************************************************
Func OpenEclipse()
Run($testCaseEclipseExePath)
WinWaitActive("Workspace Launcher")
AutoItSetOption ( "SendKeyDelay", 50)
Send($testCaseWorkSpacePath)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{TAB 3}")
Send("{Enter}")
 WinWaitActive("[Title:Java EE - Eclipse]")
 #comments-start
if $JunoOrKep = "Juno" Then
   WinWaitActive("[Title:Java EE - Eclipse]")
   Else
   WinWaitActive("[Title:Web - Eclipse]")
EndIf
#comments-end
EndFunc
;***************************************************************

;***************************************************************
;Function to create Java Project
;***************************************************************
Func CreateJavaProject()
Send("!fnd")
WinWaitActive("[Title:New Dynamic Web Project]")
AutoItSetOption ( "SendKeyDelay", 50)
Send($testCaseProjectName)
AutoItSetOption ( "SendKeyDelay", 400)
Send("!f")
WinWaitActive("[Title:Java EE - Eclipse]")
EndFunc
;***************************************************************

;***************************************************************
;Function to create JSP file and insert code
;***************************************************************
Func CreateJSPFile()
sleep(3000)
AutoItSetOption ( "SendKeyDelay", 100)
Send("{APPSKEY}")
Send("{down}")
Send("{RIGHT}")
Send("{down 14}")
Send("{enter}")
Send($testCaseJspName)
Send("!f")
Local $temp = "Java EE - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
Sleep(2000)
WinWaitActive($temp)

; Calling the Winchek Function
Local $funame, $cntrlname
$cntrlname =  "Java EE - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
$funame = "CreateJSPFile"
wincheck($funame,$cntrlname)

Send("^a")
Send("{Backspace}")
ClipPut($testCaseJspText)
Send("^v")
Send("^+s")
EndFunc
;******************************************************************

;***************************************************************
;Function to create Azure project
;***************************************************************
Func CreateAzurePackage()
WinWaitActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
Sleep(3000)
MouseClick("primary",119, 490, 1)
Send("^+{NUMPADDIV}}")
Send("{APPSKEY}")
Sleep(1000)
if $JunoOrKep = "Juno" Then
Send("g")
Else
Send("e")
EndIf
Send("{Left}{UP}{right}{Enter}")
WinWaitActive("[Title:New Azure Deployment Project]")
AutoItSetOption ( "SendKeyDelay", 50)
Send($testCaseAzureProjectName)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{TAB 3}")
Send("{Enter}")

;JDK configuration
sleep(2000)
Local $cmp = StringCompare($testCaseCheckJdk,"Check")
   if $cmp = 0 Then
	   ControlCommand("New Azure Deployment Project","","[CLASSNN:Button5]","UnCheck", "")
	   sleep(2000)
	  ControlCommand("New Azure Deployment Project","","[CLASSNN:Button5]","Check", "")
   EndIf
AutoItSetOption ( "SendKeyDelay", 100)
Send("{TAB}")
Send("+")
Send("{End}")
Send("{BACKSPACE}")
Send($testCaseJdkPath)
Send("!N")

;Server Configuration
sleep(2000)
Local $cmp = StringCompare($testCaseCheckLocalServer,"Check")
   if $cmp = 0 Then
	   ControlCommand("New Azure Deployment Project","","[CLASSNN:Button10]","UnCheck", "")
	   sleep(2000)
	  ControlCommand("New Azure Deployment Project","","[CLASSNN:Button10]","Check", "")
   EndIf
Send("{TAB}")
Send("+")
Send("{END}")
send("{BACKSPACE}")
Send($testCaseServerPath)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{TAB 2}")
for $count = $testCaseServerNo to 0 step -1
Send("{Down}")
Next
Send("!F")
EndFunc
;******************************************************************

;****************************************************************
;Function to enable SSL Offloading
;***************************************************************
Func EnableSSLOffloading()
Sleep(2000)
AutoItSetOption ( "SendKeyDelay", 100)
Send("{Up}{Enter}{down}{down}{down}{APPSKEY}")
Sleep(1000)
;if $JunoOrKep = "Juno" Then
;Send("g")
;Else
;Send("e")
;EndIf
if $JunoOrKep = "Juno" Then
Send("{s}{Left}{UP}{Right}{s 2}{Enter}")
Else
Send("{s}{Left}{UP}{UP}{Right}{s 2}{Enter}")
EndIf
WinWaitActive("[Title:Properties for WorkerRole1]")
ControlCommand("Properties for WorkerRole1","","[CLASSNN:Button1]","Check", "")
WinWaitActive("[Title:SSL Offloading]")
Send("{Enter}")
Send("{Tab 4}")
;Dim $hWnd = WinGetHandle("[Title:Properties for WorkerRole1]")
;MsgBox("","",$hWnd)
;Local $hToolBar = ControlGetHandle($hWnd, "", "[Text:<a>Certificates...</a>]")
;MsgBox("","",$hToolBar)
;ControlClick($hWnd,"",6424586,"left",1)
 MouseClick("primary",1039, 133, 1)
WinWaitActive("[Title:Properties for MyAzureProject]")
Send("{Tab 2}")
Send("{Enter}")
WinWaitActive("[Title:Certificate]")
Send("{Tab 2}")
Send("{Enter}")
Sleep(2000)
Send($testcaseCertificatePath)
;ClipPut($testcaseCertificatePath)
Send("!O")
Send("{Tab 2}")
Send("{Enter}")
Send("{Tab}{Enter}")
Send("{Tab 3}")
Send("{Enter}")
EndFunc
;*********************************************************************

;*****************************************************************
;Function to publish to cloud
;****************************************************************
Func PublishToCloud()
Sleep(2000)
WinWaitActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
Send("{Up 3}")
Send("{APPSKEY}")
Sleep(1000)
if $JunoOrKep = "Juno" Then
Send("g")
Else
Send("e")
EndIf
Send("{Left}")
Send("{UP}")
;Send("{Down 21}")
Send("{Right}")
Send("{Enter}")

WinWaitActive("Publish Wizard")
Sleep(3000)
while 1
Dim $hnd =  WinGetText("Publish Wizard","")
StringRegExp($hnd,"Loading Account Settings...",1)
Local $reg = @error
if $reg > 0 Then ExitLoop
WEnd
Send("{TAB}")

for $count = $testCaseSubscription to 1 step -1
Send("{Down}")
Next

Send("{TAB 2}")
for $count = $testCaseStorageAccount to 1 step -1
Send("{Down}")
Next

Send("{TAB 2}")
for $count = $testCaseServiceName to 1 step -1
Send("{Down}")
Next

Send("{TAB 2}")
for $count = $testCaseTargetOS to 1 step -1
Send("{Down}")
Next

Send("{TAB}")
for $count = $testCaseTargetEnvironment to 1 step -1
Send("{Down}")
Next

Send("{TAB}")
Local $cmp = StringCompare($testCaseCheckOverwrite,"UnCheck")
   if $cmp = 0 Then
	   ControlCommand("Publish Wizard","","[CLASSNN:Button4]","Check", "")
	   sleep(3000)
	  ControlCommand("Publish Wizard","","[CLASSNN:Button4]","UnCheck", "")
   Else
	  ControlCommand("Publish Wizard","","[CLASSNN:Button4]","UnCheck", "")
	   sleep(3000)
	  ControlCommand("Publish Wizard","","[CLASSNN:Button4]","Check", "")
   EndIf

Send("{TAB}")
AutoItSetOption ( "SendKeyDelay", 100)
Send($testCaseUserName)
Send("{TAB}")
Send($testCasePassword)
Send("{TAB 2}")
Send($testCasePassword)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{TAB}")
ControlCommand("Publish Wizard","","[CLASSNN:Button5]","Check", "")
Send("{TAB}")
Send("{Enter}")
EndFunc
;***************************************************************************

;*****************************************************************
;Function to check the status of RDP and Open Excel
;****************************************************************
Func CheckRDPConnection()
Local $tempTime = _Date_Time_GetLocalTime()
Local $timeDateStamp = _Date_Time_SystemTimeToDateTimeStr($tempTime)
Local $RDPWindow = ControlCommand("Remote Desktop Connection","","[CLASSNN:Button1]","IsVisible", "")
;MsgBox("","",$RDPWindow,3)
Dim $oExcel = _ExcelBookNew(0)

If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "parameter is not a number")
    Exit
EndIf

_ExcelWriteCell($oExcel, "Date And Time", 1, 1)
_ExcelWriteCell($oExcel, $timeDateStamp , 2, 1)
_ExcelWriteCell($oExcel, "RDPConnectionStatus", 1, 2)
_ExcelWriteCell($oExcel, "Test Result" , 1, 3)

if $RDPWindow = 1 Then
_ExcelWriteCell($oExcel, "Yes", 2, 2)
Send("{TAB 3}")
Send("{Enter}")
Else
_ExcelWriteCell($oExcel, "No", 2, 2)
EndIf

;Local $flag = _ExcelBookSaveAs($oExcel, @ScriptDir & "\Result" & @ScriptName, "xls",0,1)
;If $flag <> 1 Then MsgBox($MB_SYSTEMMODAL, "Not Successful", "File Not Saved!", 5)
;_ExcelBookClose($oExcel, 1, 0)
EndFunc
;***************************************************************************************

;**************************************************************************
;Function to check publish key word in Azure activity log and update excel
;**************************************************************************
Func ValidateTextAndUpdateExcel()
MouseClick("primary",565, 632, 1)
Local $string =  ControlGetText("Java EE - MyHelloWorld/WebContent/newsession.jsp - Eclipse","","[CLASS:SysLink]")
$cmp = StringRegExp($string,'<a>Published</a>',0)

;Check in webpage and update excel
Send("{TAB}")
Send("{Enter}")
Sleep(5000)
Send("{F6}")
Send("^c")
Local $url = ClipGet();
Local $temp = $url & $testCaseProjectName
Local $oIE = _IECreate($temp,1,0,1,0)
_IELoadWait($oIE)
Local $readHTML = _IEBodyReadText($oIE)
Local $iCmp = StringRegExp($readHTML,$testCaseValidationText,0)

;Check for error
If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "File does not exist")
    Exit
 EndIf

if $iCmp = 1 Then
;MsgBox ($MB_SYSTEMMODAL, "Test Result", "Test Passed")
_ExcelWriteCell($oExcel, "Test Passed" , 2, 3)
Else
;MsgBox ($MB_SYSTEMMODAL, "Test Result", "Test Failed")
_ExcelWriteCell($oExcel, "Test Failed" , 2, 3)
EndIf

Local $flag = _ExcelBookSaveAs($oExcel, @ScriptDir & "\" & $testCaseName & "Result", "xls",0,1)
If Not @error Then MsgBox($MB_SYSTEMMODAL, "Success", "File was Saved!", 3)
_ExcelBookClose($oExcel, 1, 0)
EndFunc
;*******************************************************************************


Func Delete()
Dim $hWnd = WinGetHandle("[CLASS:SWT_Window0]")
Local $hToolBar = ControlGetHandle($hWnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
WinActivate($hToolBar)
MouseClick("primary",119, 490, 1)
Send("^+{NUMPADDIV}}")
for $i = 6 to 1 Step - 1
   Local $chk = _GUICtrlTreeView_GetCount($hToolBar)

if $chk = 0 Then
	ExitLoop
 Else
	MouseClick("primary",119, 490, 1)
   Send("^+{NUMPADDIV}}")
   Send("{RIGHT}")
   Send("{DOWN}")
   Send("{UP}")
   Send("{DELETE}")
   Send("{SPACE}")
   Send("{ENTER}")

   EndIf
Next
EndFunc