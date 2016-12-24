
' ====================================================================== '
'
' Group Membership Network Drive Mapper
'
' Copyright (C) 2016 tag.
'
' ====================================================================== '

' Prepare environment values
strScriptDirPath = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - Len(WScript.ScriptName))
strNetwkDrvMapFilePath = "GrpMemDrvMapList.txt"
strDelimiter = ","

' Prepare objects
Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
Dim objStream : Set objStream = objFso.OpenTextFile(objFso.BuildPath(strScriptDirPath, strNetwkDrvMapFilePath))
Dim objNetwk : Set objNetwk = WScript.CreateObject("WScript.Network")
Dim objShell : Set objShell = WScript.CreateObject("WScript.Shell")
Dim objAdSysInfo : Set objAdSysInfo = CreateObject("ADSystemInfo")
Dim objUser : Set objUser = GetObject("LDAP://" & objAdSysInfo.UserName)
Dim objApp : Set objApp = CreateObject("Shell.Application")

' ---------------------------------------------------------------------- '

' Get group memberships
Dim objGroup
strBelongGroupList = GetPrimaryGroup(objUser)	' Primary Group
For Each strGroup In objUser.memberOf	' Secondary Group, Tertiary Group, ...
	Set objGroup = GetObject("LDAP://" & strGroup)
	strGroupName = objGroup.name
	strBelongGroupList = strBelongGroupList & strDelimiter & Right(strGroupName, Len(strGroupName) - 3)
Next
'WScript.Echo strBelongGroupList

' ---------------------------------------------------------------------- '

' Read the mapping file
Do Until objStream.AtEndOfStream
	strLine = objStream.ReadLine
	'WScript.Echo strLine
	If InStr(strLine, "#") <> 1 Then	' If not comment line
		aryDrvMapList = Split(strLine, vbTab)
		For Each strBelongGroup In Split(strBelongGroupList, strDelimiter)
			If aryDrvMapList(0) = strBelongGroup Then
				strLocalName = aryDrvMapList(1)	' Drive Letter
				strRemoteName = aryDrvMapList(2)	' Network Place
				strDriveName = aryDrvMapList(3)	' Drive Name

				' Unmount if already mounted
				If objFso.DriveExists(strLocalName) Then
					objNetwk.RemoveNetworkDrive(strLocalName)
				End If

				' Mount
				Call objNetwk.MapNetworkDrive(strLocalName, strRemoteName)

				' Rename
				objApp.NameSpace(strLocalName).Items().Item().Name = strDriveName
			End If
		Next
	End If
Loop

WScript.Quit

' ---------------------------------------------------------------------- '
'
' Functions
'
' ---------------------------------------------------------------------- '

' Get PrimaryGroup from AD object
Function GetPrimaryGroup(objADs)
	objVal = objADs.ObjectSid
	strSID = ""
	For i = lBound(objVal) to uBound(objVal) - 4
		hByte = ascb(midb(objVal, i+1, 1))
		lByte = hByte mod 16
		hByte = hByte \ 16
		strSID = strSID & hex(hByte) & hex(lByte)
	Next

	strRid = right("0000000" & hex(objADs.primaryGroupID), 8)
	hRid = right(left(strRid,4), 2) & left(left(strRid,4), 2)
	lRid = right(right(strRid,4), 2) & left(right(strRid,4), 2)
	strSID = strSID & lRid & hRid

	Set objGroup = GetObject("LDAP://<SID=" & strSID & ">")

	Dim objPrimaryGroup : Set objPrimaryGroup = GetObject("LDAP://" & objGroup.distinguishedName)
	GetPrimaryGroup = Right(objPrimaryGroup.name, Len(objPrimaryGroup.name) - 3)
End Function
