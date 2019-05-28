'cscript SearchAD.vbs >output.txt
'user_info.vbs

' Usage:
'       cscript //Nologo user_info.vbs

' List User properties as displayed in ADUC

On Error Resume Next
Dim objSysInfo, objUser
Set objSysInfo = CreateObject("ADSystemInfo")

' Currently logged in User
Set objUser = GetObject("LDAP://" & objSysInfo.UserName)
WScript.Echo "#1 " & objUser
WScript.Echo "#2 " & objSysInfo.UserName
 ' or specific user:
'Set objUser = GetObject("LDAP://CN=20067,OU=Students,DC=SCHOOL,DC=HCPS")

WScript.Echo "DN: " & objUser.distinguishedName

WScript.Echo ""
WScript.Echo "GENERAL"
'WScript.Echo "First name: " & objUser.givenName
WScript.Echo "First name: " & objUser.FirstName
'WScript.Echo "Last name: " & objUser.sn
WScript.Echo "Last name: " & objUser.LastName
WScript.Echo "Display name: " & objUser.displayName
'WScript.Echo "Display name: " & objUser.FullName

WScript.Echo ""
WScript.Echo "ACCOUNT"
WScript.Echo "User logon name: " & objUser.userPrincipalName
WScript.Echo "AccountDisabled: " & objUser.AccountDisabled
' WScript.Echo "Account Control #: " & objUser.userAccountControl
WScript.Echo "Logon Hours: " & objUser.logonHours
WScript.Echo "Logon On To (Logon Workstations): " & objUser.userWorkstations
' WScript.Echo "User must change password at next logon: " & objUser.pwdLastSet
' WScript.Echo "Account expires end of (date): " & objUser.accountExpires
WScript.Echo ""
WScript.Echo "PROFILE"
WScript.Echo "Profile path: " & objUser.profilePath
' WScript.Echo "Profile path: " & objUser.Profile
WScript.Echo "Logon script: " & objUser.scriptPath
WScript.Echo "Home folder, local path: " & objUser.homeDirectory
WScript.Echo "Home folder, Connect, Drive: " & objUser.homeDrive
WScript.Echo "Home folder, Connect, To:: " & objUser.homeDirectory
