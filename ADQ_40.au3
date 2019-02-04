#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=C:\Users\carmine.santoro\Pictures\Sirea-Angry-Birds-Bird-blue.ico
#AutoIt3Wrapper_Outfile_x64=ADHelp2.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=AD helper
#AutoIt3Wrapper_Res_Fileversion=3.0.0.3
#AutoIt3Wrapper_Res_ProductVersion=3.0.0.3
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <GUIConstants.au3>
#include <Misc.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <ColorConstants.au3>
#include <GUIConstantsEx.au3>
#include <Date.au3>

Const $ADS_NAME_INITTYPE_GC = 3
Const $ADS_NAME_TYPE_NT4 = 3
Const $ADS_NAME_TYPE_1779 = 1
Dim $unlock
Dim $mgrvalue
Dim $mgrsplit
Dim $manager
Dim $mgr
Dim $title
Dim $pwdexpires
$oMyError = ObjEvent("AutoIt.Error", "ComError")
$objRootDSE = ObjGet("LDAP://RootDSE")
$username = InputBox("Username", "Please input a username:", "please enter username or email address")
If StringInStr($username, "@") Then
	$aa = StringSplit($username, "@")
	$username = $aa[1]
ElseIf StringInStr($username, "\") Then

	$aa = StringSplit($username, "\")
	$username = $aa[2]
ElseIf $username = "" Then
	Exit
EndIf

If @error Then
	MsgBox(0, 'username', 'Username does not exist or not able to communicate with ' & @LogonDomain)
Else
	; DNS domain name.
	$objTrans = ObjCreate("NameTranslate")
	$objTrans.Init($ADS_NAME_INITTYPE_GC, "")
	$objTrans.Set($ADS_NAME_TYPE_1779, @LogonDomain)
	$objTrans.Set($ADS_NAME_TYPE_NT4, @LogonDomain & "\" & $username)
	$strUserDN = $objTrans.Get($ADS_NAME_TYPE_1779)
	$UserObj = ObjGet("LDAP://" & $strUserDN)
	If @error Then
		MsgBox(0, 'username', 'Username does not exist or not able to communicate with ' & @LogonDomain)
	Else
		;MsgBox(0, 'test', 'test:  ' & $test)
		Call("Displayinfo")


	EndIf
EndIf
$UserObj = ""
$oMyError = ObjEvent("AutoIt.Error", "")
;COM Error function
Func ComError()
	If IsObj($oMyError) Then
		$HexNumber = Hex($oMyError.number, 8)
		SetError($HexNumber)
	Else
		SetError(1)
	EndIf
	Return 0
EndFunc   ;==>ComError



Func Displayinfo()

	If Not FileExists(@ScriptDir & "\ADLook") Then
		DirCreate(@ScriptDir & "\ADLook")
	EndIf

	If Not FileExists(@ScriptDir & "\ADLook" & "\log.txt") Then
		_FileCreate(@ScriptDir & "\ADLook" & "\log.txt")
	EndIf

	FileWriteLine(@ScriptDir & "\ADLook" & "\log.txt", $username)

	GUICreate("Active Directory Information", 500, 620, 300, 300)
	GUICtrlCreateLabel("Username: ", 10, 10, 60, 20)
	GUICtrlCreateLabel("First Name: ", 10, 30, 60, 20)
	GUICtrlCreateLabel("Last Name: ", 200, 30, 60, 20)
	GUICtrlCreateLabel("Display Name: ", 10, 50, 100, 20)
	GUICtrlCreateLabel("Title: ", 10, 70, 100, 20)
	GUICtrlCreateLabel("Manager: ", 10, 90, 100, 20)
	GUICtrlCreateLabel("Description: ", 10, 130, 100, 20)
	GUICtrlCreateLabel("Office: ", 10, 170, 60, 20)
	GUICtrlCreateLabel("Department: ", 10, 215, 100, 20)
	GUICtrlCreateLabel("Telephone Number: ", 10, 260, 90, 40)
	GUICtrlCreateLabel("Mobile Number: ", 10, 300, 100, 20)
	GUICtrlCreateLabel("Home Number: ", 10, 330, 100, 20)
	GUICtrlCreateLabel("Email Address: ", 10, 350, 100, 20)
	GUICtrlCreateLabel("Employee ID: ", 10, 375, 100, 20)
	GUICtrlCreateLabel("Logon Script: ", 10, 410, 100, 20)
	GUICtrlCreateLabel("Account:", 10, 430, 100, 20)
	;GUICtrlCreateLabel("Number of bad logon attempts since last reset: ", 310, 420, 120, 40)
	GUICtrlCreateLabel("Password Last Changed: ", 10, 460, 120, 40)
	GUICtrlCreateLabel("90 Day Password Expiration: ", 10, 490, 190, 40)
	GUICtrlCreateLabel("Last Logon: ", 10, 520, 100, 20)
	;GUICtrlCreateLabel("Last bad Password: ", 10, 550, 100, 20)
	GUICtrlCreateLabel("Account Created on: ", 10, 550, 100, 20)

	$font = "Tahoma"
	GUISetFont(9, 600, $font) ; will display underlined characters
	;$unlock = GUICtrlCreateButton("UNLOCK Account", 180, 425, 120, 25)
	;GUICtrlSetState($unlock, $Gui_Disable)
	GUICtrlCreateEdit('' & $username, 100, 10, 200, 20)
	GUICtrlSetColor(-1, 0x0000CC) ; Blue
	GUICtrlCreateLabel('' & $UserObj.FirstName, 100, 30, 100, 20)
	GUICtrlCreateLabel('' & $UserObj.LastName, 300, 30, 150, 50)
	GUICtrlCreateEdit('' & $UserObj.FullName, 100, 50, 300, 20)
	GUICtrlCreateLabel('' & $UserObj.Title, 100, 70, 100, 20)
	$title = GUICtrlRead($title)
	If $title = 0 Then
		GUICtrlCreateLabel('', 100, 70, 100, 20)
	EndIf

	$mgr = GUICtrlCreateLabel('' & $UserObj.Manager, 100, 90, 400, 70)
	$mgrvalue = GUICtrlRead($mgr)
	$mgrsplit = StringSplit("" & $mgrvalue, ",")
	$manager = StringTrimLeft('' & $mgrsplit[1], 3)
	GUICtrlCreateLabel('' & $manager, 100, 90, 400, 70)
	GUICtrlCreateLabel('' & $UserObj.Description, 100, 130, 300, 40)
	GUICtrlCreateLabel('' & $UserObj.physicalDeliveryOfficeName, 100, 170, 100, 50)
	GUICtrlCreateLabel('' & $UserObj.Department, 100, 215, 290, 20)
	GUICtrlCreateEdit('' & $UserObj.TelephoneNumber, 100, 260, 250, 20)
	GUICtrlCreateLabel('' & $UserObj.TelephoneMobile, 100, 300, 250, 20)
	GUICtrlCreateLabel('' & $UserObj.TelephoneHome, 120, 330, 250, 20)
	GUICtrlCreateEdit('' & $UserObj.EmailAddress, 100, 350, 300, 20)
	GUICtrlCreateLabel('' & $UserObj.LoginScript, 100, 410, 200, 15)
	GUICtrlCreateEdit('' & $UserObj.employeeID, 100, 375, 200, 20)

	#cs
		$locked = GUICtrlCreateLabel("" & $UserObj.IsAccountLocked, 100, 430, 10, 20)
		If GUICtrlRead($locked) = 0 Or 39 Then
		GUICtrlCreateLabel("NOT Locked", 100, 430, 80, 15)
		GUICtrlSetBkColor(-1, 0x00ff00);Green
		Else
		MsgBox(0, 'INFO', "User Account Lock value is: " & $locked)
		GUICtrlCreateLabel("LOCKED", 10, 430, 60, 15)
		GUICtrlSetBkColor(-1, 0xff0000) ; Red
		GUICtrlSetState($unlock, $Gui_Enable)

		EndIf
	#ce

	$lastchange = $UserObj.PasswordLastChanged
	$res1 = StringRegExpReplace($lastchange, '(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})', "$1/$2/$3 $4:$5:$6")
	;MsgBox(0,"", $res1)
	;$Date = StringMid($lastchange, 5, 2) & "/" & StringMid($lastchange, 7, 2) & "/" & StringMid($lastchange, 1, 4)
	;$Time = StringMid($lastchange, 9, 2) & ":" & StringMid($lastchange, 11, 2) & ":" & StringMid($lastchange, 13, 2)
	;GUICtrlCreateLabel($Date & " " & $Time, 100, 460, 150, 20)
	GUICtrlCreateLabel($res1, 135, 460, 150, 20)
	;$pwdexpires = StringMid($lastchange, 5, 2) + 3 & "/" & (StringMid($lastchange, 7, 2) - 2) & "/" & StringMid($lastchange, 1, 4)
	; added a -2 to subtract one day from the expiration date as I found that it is one day less than what is in AD.
	;GUICtrlCreateLabel($pwdexpires & ' ' & $Time, 100, 490, 150, 20)
	$strPasswordExpirationDate = $UserObj.PasswordExpirationDate
	$res = StringRegExpReplace($strPasswordExpirationDate, '(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})', "$1/$2/$3 $4:$5:$6")
	;MsgBox(0, "", $res)

	Local $iDateCalc = _DateDiff('D', $res1, _NowCalc())

	If $iDateCalc >= 90 Then
		GUICtrlCreateLabel($pwdexpires & ' ' & $res, 150, 490, 150, 20)
		GUICtrlSetColor(-1, $COLOR_RED)
	ElseIf $iDateCalc >= 80 Then
		GUICtrlCreateLabel($pwdexpires & ' ' & $res, 150, 490, 150, 20)
		GUICtrlSetColor(-1, 0xe65c00)
	ElseIf $iDateCalc <= 10 Then
		GUICtrlCreateLabel($pwdexpires & ' ' & $res, 150, 490, 150, 20)
		GUICtrlSetColor(-1, 0x009933)
	Else
		GUICtrlCreateLabel($pwdexpires & ' ' & $res, 150, 490, 150, 20)
	EndIf


	; Account create date
	$create = $UserObj.WhenCreated
	$Date = StringMid($create, 5, 2) & "/" & StringMid($create, 7, 2) & "/" & StringMid($create, 1, 4)
	$Time = StringMid($create, 9, 2) & ":" & StringMid($create, 11, 2) & ":" & StringMid($create, 13, 2)
	GUICtrlCreateLabel($Date & " " & $Time, 120, 550, 150, 20)
	$create1 = StringMid($create, 5, 2) + 3 & "/" & StringMid($create, 7, 2) & "/" & StringMid($create, 1, 4)

	$lastlogin = $UserObj.LastLogin

	$Date = StringMid($lastlogin, 5, 2) & "/" & StringMid($lastlogin, 7, 2) & "/" & StringMid($lastlogin, 1, 4)
	$Time = StringMid($lastlogin, 9, 2) & ":" & StringMid($lastlogin, 11, 2) & ":" & StringMid($lastlogin, 13, 2)
	GUICtrlCreateLabel($Date & " " & $Time, 100, 520, 150, 20)
	GUICtrlCreateLabel("Bad Password Count: " & $UserObj.badPwdCount, 10, 580, 170, 20)


	;MsgBox(0,"", $UserObj.badPwdCount)
	;$UserObj.homeDirectory = full unc path
	;$UserObj.homedrive = Drive letter
	;$UserObj.badPwdCount = counts in incriments of 2 if logging onto blackboard, 1 if computer.

	;GUICtrlCreateLabel($UserObj.badPasswordTime, 100, 550, 150, 20)

	#cs
		$badlogin = GUICtrlCreateLabel("" & $UserObj.BadLoginCount, 430, 430, 20, 15)
		If GUICtrlRead($badlogin) = 0 Then
		GUICtrlSetBkColor(-1, 0x00ff00);Green
		Else
		GUICtrlSetBkColor(-1, 0xff0000) ; Red
		EndIf
	#ce

	GUISetState()



	While 1
		$msg = GUIGetMsg()
		Select
			Case $msg = $unlock
				;If $UserObj.IsAccountLocked Then
				;$UserObj.IsAccountLocked = False
				;$UserObj.SetInfo
				;MsgBox(0, 'INFO', "User Account was Unlocked. It will take approximately 5 mins to reflect this change.")
				;GUICtrlCreateLabel("" & $UserObj.IsAccountLocked, 100, 430, 10, 20)
				;EndIf

			Case $msg = $GUI_EVENT_CLOSE
				Exit
		EndSelect
	WEnd



EndFunc   ;==>Displayinfo

; http://www.autoitscript.com/forum/topic/28404-active-directory-helper/?hl=%2Bprofile+%2Bmanager
; https://msdn.microsoft.com/en-us/library/ms675244(v=vs.85).aspx
; https://www.autoitscript.com/forum/topic/135921-query-ldap-for-password-expiration/
