'cscript SearchAD.vbs >output.txt

'SearchAD.vbs
On Error Resume Next

' Connect to the LDAP server's root object
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")
strTarget = "LDAP://" & strDNSDomain
wscript.Echo "Starting search from " & strTarget

' Connect to Ad Provider
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCmd =   CreateObject("ADODB.Command")
Set objCmd.ActiveConnection = objConnection 

' Show only computers
'objCmd.CommandText = "SELECT Name, ADsPath FROM '" & strTarget & "' WHERE objectCategory = 'computer'"

' or show only users
objCmd.CommandText = "SELECT Name, ADsPath FROM '" & strTarget & "' WHERE objectCategory = 'user'"

' or show only groups
'objCmd.CommandText = "SELECT Name, ADsPath FROM '" & strTarget & "' WHERE objectCategory = 'group'"

Const ADS_SCOPE_SUBTREE = 2
objCmd.Properties("Page Size") = 100
objCmd.Properties("Timeout") = 30
objCmd.Properties("Searchscope") = ADS_SCOPE_SUBTREE
objCmd.Properties("Cache Results") = False

Set objRecordSet = objCmd.Execute

' Iterate through the results
objRecordSet.MoveFirst
Do Until objRecordSet.EOF
   sComputerName = objRecordSet.Fields("Name")
   wscript.Echo sComputerName
   objRecordSet.MoveNext
Loop