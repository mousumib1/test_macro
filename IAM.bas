Attribute VB_Name = "Create_IAM"

Public Sub CreatePropertyFile_IAM(jsonpath As String)
'*****************************************************
' Purpose:   Create property file for IAM as an xml format.
' Notes:     Requires an IAM parameter sheet.
'
' Date:       Description:
' 04/03/15    Created.
'*****************************************************

Dim relativeFilePath As String
Dim i As Long
Dim intFileNum As Integer
Dim appPath As String
Dim outputFile As String
Dim findText As String
Dim groupHeaderText As String
Dim userHeaderText As String
Dim roleHeaderText As String
Dim passwordPolicyHeaderText As String
Dim groupColumnName As String
Dim userColumnName As String
Dim userGroupColumnName As String
Dim passwordPolicyColumnName As String
Dim passwordPolicyColumn2Name As String

'***************************************************
' Any changes in of the following value in the IAM parameter sheet,
' please modify with the same in the respective value of following variable.

groupHeaderText = "1.1?Groups"
userHeaderText = "1.2?Users"
roleHeaderText = "1.3?Roles"
passwordPolicyHeaderText = "1.4?PasswordPolicy"

groupColumnName = "Group Name"
userColumnName = "User Name"
userGroupColumnName = "Group"
passwordPolicyColumnName = "Item"
passwordPolicySettingColumnName = "Setting"

'*****************************************************

'***************************************************
' Please change the path of the following of the output XML file

appPath = jsonpath 
' "F:\aws_work\awscoe\APRIL2015\poc\macro_attach\"


'***************************************************

relativeFilePath = "iam_properties.xml"


'Build output file
outputFile = appPath + relativeFilePath
intFileNum = 1

'Open the file and write output
Open outputFile For Output As #intFileNum

Print #intFileNum, "<?xml version=""1.0"" encoding=""utf-8""?>"


' Start: This section is used to retrieve the groups
Dim groupRowCount As Integer
Dim isGroupsFound As Boolean
Dim isGroupHeaderFound As Boolean
Dim groupColumnNo As Integer


findText = groupHeaderText
isGroupsFound = False
isGroupHeaderFound = False

'Find the end of group section
groupRowCount = Rows.Find(What:=userHeaderText, LookIn:=xlValues).Row
Print #intFileNum, "<GROUPS>"

For i = 1 To groupRowCount

    If Not isGroupHeaderFound Then
        Set findRow = Rows(i).Cells.Find(What:=findText, LookIn:=xlValues)
    ElseIf Not IsEmpty(Cells(i, groupColumnNo)) Then
        Print #intFileNum, " <GROUP>"
        Print #intFileNum, "     <GROUPNAME>" & Cells(i, groupColumnNo) & "</GROUPNAME>"
        Print #intFileNum, " </GROUP>"
    End If

    If Not findRow Is Nothing Then
       If Not isGroupsFound Then
            findText = groupColumnName
            isGroupsFound = True
        ElseIf findText = groupColumnName And Not isGroupHeaderFound Then
            groupColumnNo = findRow.Column
            isGroupHeaderFound = True
       End If
    End If
    
Next i

Print #intFileNum, "</GROUPS>"

' Start: This section is used to retrieve the users

Dim userRowCount As Integer
Dim isUsersFound As Boolean
Dim isUserHeaderFound As Boolean
Dim userColumnNo As Integer

findText = userHeaderText
isUsersFound = False
isUserHeaderFound = False

'userRowCount = Rows.Find(What:=roleHeaderText, LookIn:=xlValues).Row
Print #intFileNum, "<USERS>"

For i = groupRowCount To Range("A1").SpecialCells(xlCellTypeLastCell).Row

    If Not isUserHeaderFound Then
        Set findRow = Rows(i).Cells.Find(What:=findText, LookIn:=xlValues)
    ElseIf Not IsEmpty(Cells(i, userColumnNo)) Then
        Print #intFileNum, " <USER>"
        Print #intFileNum, "     <USERNAME>" & Cells(i, userColumnNo) & "</USERNAME>"
        If Not IsEmpty(Cells(i, groupColumnNo)) Then
            If Cells(i, groupColumnNo) = "-" Then
                Print #intFileNum, "     <GROUPNAME></GROUPNAME>"
            Else
                Print #intFileNum, "     <GROUPNAME>" & Cells(i, groupColumnNo) & "</GROUPNAME>"
            End If
            
        End If
        Print #intFileNum, " </USER>"
    End If

    If Not findRow Is Nothing Then
       If Not isUsersFound Then
            findText = userColumnName
            isUsersFound = True
        ElseIf findText = userColumnName And Not isUserHeaderFound Then
            userColumnNo = findRow.Column
            groupColumnNo = Rows(i).Cells.Find(What:=userGroupColumnName, LookIn:=xlValues).Column
            isUserHeaderFound = True
       End If
    End If
    
Next i

Print #intFileNum, "</USERS>"


' Start: This section is used to retrieve the password policy

'Dim passwordPolicyRowCount As Integer
'Dim isPasswordPolicyFound As Boolean
'Dim isPasswordPolicyHeaderFound As Boolean
'Dim itemColumnNo As Integer
'Dim settingColumnNo As Integer

'findText = passwordPolicyHeaderText
'isPasswordPolicyFound = False
'isUserHeaderFound = False


'For i = userRowCount To Range("A1").SpecialCells(xlCellTypeLastCell).Row

    'If Not isPasswordPolicyHeaderFound Then
        'Set findRow = Rows(i).Cells.Find(What:=findText, LookIn:=xlValues)
    'ElseIf Not IsEmpty(Cells(i, itemColumnNo)) Then
        
        'If Not IsEmpty(Cells(i, settingColumnNo)) And Cells(i, settingColumnNo) <> "Default" Then
            'Print #intFileNum , Cells(i, itemColumnNo)
            'Print #intFileNum , Cells(i, settingColumnNo)
        'End If
   ' End If

    'If Not findRow Is Nothing Then
       'If Not isPasswordPolicyFound Then
            'findText = passwordPolicyColumnName
            'isPasswordPolicyFound = True
        'ElseIf findText = passwordPolicyColumnName And Not isPasswordPolicyHeaderFound Then
            'itemColumnNo = findRow.Column
            'Print #intFileNum , Cells(i, itemColumnNo)
            'settingColumnNo = Rows(i).Cells.Find(What:=passwordPolicySettingColumnName, LookIn:=xlValues).Column
            'isPasswordPolicyHeaderFound = True
       'End If
    'End If
'Next i



'Close the output file
Close #intFileNum
End Sub
		