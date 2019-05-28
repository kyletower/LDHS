' Programmer: Kyle Tower
' Date: 11.01.2018
' Description: Creates shortcuts to each of my student's directory
' so that I can grade programming assignments.
' Source(s): ss64.com

On Error Resume Next

' Declare variables
Dim objSysInfo, objUser

Set objSysInfo = CreateObject("ADSystemInfo")

' Declare variables for each block's lunch IDs
Dim A1_lunch_ID, A2_lunch_ID, A3_lunch_ID, A4_lunchID
Dim B1_lunch_ID, B2_lunch_ID, B3_lunch_ID, B4_lunchID

' Initialize each array of lunch IDs
' A1_lunch_ID = Array("#####", "#####", ... )
' A2_lunch_ID = Array("41217", "41797", "26020", "19815", "56781", "31593", "20733", "60999", "30595", "32015", "26093", "47105")
' A3_lunch_ID = Array("25267", "46876", "30606", "28217", "23571", "43570", "41743", "51293", "22704", "43441", "33291", "44198", "36096", "35830", "36569", "41522", "65553", "43940", "30934", "33456", "34938", "41303")
'A4_lunch_ID = Array("43451", "44594", "48895", "51351", "46347", "47157", "52298", "46940", "43935", "43746", "63688", "43763", "18987", "46684", "36776", "46693")

' B1_lunch_ID = Array("43027", "42444", "54646", "43357", "20906", "54842", "43876", "45163", "28310", "61253", "46860", "31252", "25975", "41524", "41525", "26462", "41121", "41809", "51226", "31241", "44097", "45382", "33793", "55974", "23531")
'B2_lunch_ID = Array("41972", "54341", "45729", "46304", "58251", "67310", "30387", "46557", "67537", "30286", "46715", "41228", "46689", "36544", "43749", "42940", "58208")
'B3_lunch_ID = Array("46876", "42444", "28217", "43357", "43876", "43587", "51293", "46860", "33291", "36096", "41522", "41524", "45127", "43846", "30934", "30205", "46147")
B4_lunch_ID = Array("45814", "46526", "44594", "51270", "46556", "46715", "67628", "67630", "63109", "67587")

Dim oFSO
Set oFSO = CreateObject("Scripting.FileSystemObject")

' -------------------------- BLOCK A2 -------------------------- '
' Create a new folder for Block A2
oFSO.CreateFolder "A2"

' Iterate through each student ID
For student = 0 To UBound(A2_lunch_ID)
    ' Get specific user:
    Set objUser = GetObject("LDAP://CN=" & A2_lunch_ID(student) & ",OU=Lee Davis Students,OU=Students,DC=SCHOOL,DC=HCPS")

    ' Split full name of user (using a blank space as delimiter) into an array
    fullName = Split(objUser.displayName, " ")


    Set oWS = WScript.CreateObject("WScript.Shell")

    ' Create shortcut whereby lastName = fullName(1) and firstName = fullName(0)
    sLinkFile = "A2\" & fullName(1) & ", " & fullName(0) & ".LNK"
    Set oLink = oWS.CreateShortcut(sLinkFile)

    ' Path of student's drive
    oLink.TargetPath = "\\ldhs\students\" & A2_lunch_ID(student)

    ' Save the shortcut
    oLink.Save
Next

' -------------------------- BLOCK A3 -------------------------- '
' Create a new folder for Block A3
oFSO.CreateFolder "A3"

' Iterate through each student ID
For student = 0 To UBound(A3_lunch_ID)
    ' Get specific user:
    Set objUser = GetObject("LDAP://CN=" & A3_lunch_ID(student) & ",OU=Lee Davis Students,OU=Students,DC=SCHOOL,DC=HCPS")

    ' Split full name of user (using a blank space as delimiter) into an array
    fullName = Split(objUser.displayName, " ")


    Set oWS = WScript.CreateObject("WScript.Shell")

    ' Create shortcut whereby lastName = fullName(1) and firstName = fullName(0)
    sLinkFile = "A3\" & fullName(1) & ", " & fullName(0) & ".LNK"
    Set oLink = oWS.CreateShortcut(sLinkFile)

    ' Path of student's drive
    oLink.TargetPath = "\\ldhs\students\" & A3_lunch_ID(student)

    ' Save the shortcut
    oLink.Save
Next

' -------------------------- BLOCK A4 -------------------------- '
' Create a new folder for Block A4
oFSO.CreateFolder "A4"

' Iterate through each student ID
For student = 0 To UBound(A4_lunch_ID)
    ' Get specific user:
    Set objUser = GetObject("LDAP://CN=" & A4_lunch_ID(student) & ",OU=Lee Davis Students,OU=Students,DC=SCHOOL,DC=HCPS")

    ' Split full name of user (using a blank space as delimiter) into an array
    fullName = Split(objUser.displayName, " ")


    Set oWS = WScript.CreateObject("WScript.Shell")

    ' Create shortcut whereby lastName = fullName(1) and firstName = fullName(0)
    sLinkFile = "A4\" & fullName(1) & ", " & fullName(0) & ".LNK"
    Set oLink = oWS.CreateShortcut(sLinkFile)

    ' Path of student's drive
    oLink.TargetPath = "\\ldhs\students\" & A4_lunch_ID(student)

    ' Save the shortcut
    oLink.Save
Next

' -------------------------- BLOCK B1 -------------------------- '
' Create a new folder for Block B1
oFSO.CreateFolder "B1"

' Iterate through each student ID
For student = 0 To UBound(B1_lunch_ID)
    ' Get specific user:
    Set objUser = GetObject("LDAP://CN=" & B1_lunch_ID(student) & ",OU=Lee Davis Students,OU=Students,DC=SCHOOL,DC=HCPS")

    ' Split full name of user (using a blank space as delimiter) into an array
    fullName = Split(objUser.displayName, " ")


    Set oWS = WScript.CreateObject("WScript.Shell")

    ' Create shortcut whereby lastName = fullName(1) and firstName = fullName(0)
    sLinkFile = "B1\" & fullName(1) & ", " & fullName(0) & ".LNK"
    Set oLink = oWS.CreateShortcut(sLinkFile)

    ' Path of student's drive
    oLink.TargetPath = "\\ldhs\students\" & B1_lunch_ID(student)

    ' Save the shortcut
    oLink.Save
Next


' -------------------------- BLOCK B2 -------------------------- '
' Create a new folder for Block B2
oFSO.CreateFolder "B2"

' Iterate through each student ID
For student = 0 To UBound(B2_lunch_ID)
    ' Get specific user:
    Set objUser = GetObject("LDAP://CN=" & B2_lunch_ID(student) & ",OU=Lee Davis Students,OU=Students,DC=SCHOOL,DC=HCPS")

    ' Split full name of user (using a blank space as delimiter) into an array
    fullName = Split(objUser.displayName, " ")


    Set oWS = WScript.CreateObject("WScript.Shell")

    ' Create shortcut whereby lastName = fullName(1) and firstName = fullName(0)
    sLinkFile = "B2\" & fullName(1) & ", " & fullName(0) & ".LNK"
    Set oLink = oWS.CreateShortcut(sLinkFile)

    ' Path of student's drive
    oLink.TargetPath = "\\ldhs\students\" & B2_lunch_ID(student)

    ' Save the shortcut
    oLink.Save
Next


' -------------------------- BLOCK B3 -------------------------- '
' Create a new folder for Block B3
oFSO.CreateFolder "B3"

' Iterate through each student ID
For student = 0 To UBound(B3_lunch_ID)
    ' Get specific user:
    Set objUser = GetObject("LDAP://CN=" & B3_lunch_ID(student) & ",OU=Lee Davis Students,OU=Students,DC=SCHOOL,DC=HCPS")

    ' Split full name of user (using a blank space as delimiter) into an array
    fullName = Split(objUser.displayName, " ")


    Set oWS = WScript.CreateObject("WScript.Shell")

    ' Create shortcut whereby lastName = fullName(1) and firstName = fullName(0)
    sLinkFile = "B3\" & fullName(1) & ", " & fullName(0) & ".LNK"
    Set oLink = oWS.CreateShortcut(sLinkFile)

    ' Path of student's drive
    oLink.TargetPath = "\\ldhs\students\" & B3_lunch_ID(student)

    ' Save the shortcut
    oLink.Save
Next



' -------------------------- BLOCK B4 -------------------------- '
' Create a new folder for Block B4
oFSO.CreateFolder "B4"

' Iterate through each student ID
For student = 0 To UBound(B4_lunch_ID)
    ' Get specific user:
    Set objUser = GetObject("LDAP://CN=" & B4_lunch_ID(student) & ",OU=Lee Davis Students,OU=Students,DC=SCHOOL,DC=HCPS")

    ' Split full name of user (using a blank space as delimiter) into an array
    fullName = Split(objUser.displayName, " ")


    Set oWS = WScript.CreateObject("WScript.Shell")

    ' Create shortcut whereby lastName = fullName(1) and firstName = fullName(0)
    sLinkFile = "B4\" & fullName(1) & ", " & fullName(0) & ".LNK"
    Set oLink = oWS.CreateShortcut(sLinkFile)

    ' Path of student's drive
    oLink.TargetPath = "\\ldhs\students\" & B4_lunch_ID(student)

    ' Save the shortcut
    oLink.Save
Next