Dim CurrentSection As String
Dim CurrentTestCaseProperty As String

Private Type TestCase
    slogen As String
    tag As String
    priority As String
    environment As String
    execute_type As String
    test_type As String
    requirement As String
    reference As String
    description As String
    precondition As String
    action As String
    result As String
End Type

Dim CurrentTestCase As TestCase

Sub GetTestCaseSlogan()
    If Selection.Bookmarks.Exists("\Headinglevel") Then
        With Selection.Bookmarks("\Headinglevel").Range
            CurrentSection = .ListParagraphs(1).Range.ListFormat.ListString
            CurrentTestCase.slogen = CleanString(.ListParagraphs(1).Range.Text)
        End With
    End If
End Sub

Function IsSameTestCase() As Boolean
    If Selection.Bookmarks.Exists("\Headinglevel") Then
        With Selection.Bookmarks("\Headinglevel").Range
        If .ListParagraphs(1).Range.ListFormat.ListString = CurrentSection Then
            IsSameTestCase = True
            Exit Function
        Else
            IsSameTestCase = False
            Exit Function
        End If
        End With
    Else
        IsSameTestCase = False
    End If
End Function

Function GetPropertyName$()
    Dim rTmp1 As Range
    Dim rtmp2 As Range
    Dim propertyName As String
    Set rTmp1 = Selection.Range
    Set rtmp2 = Selection.Range
    While rTmp1.ListParagraphs.Count = 0
        rTmp1.MoveStart unit:=wdParagraph, Count:=-1
    Wend
    rtmp2.Select
    GetPropertyName = rTmp1.ListParagraphs(1).Range.ListFormat.ListString
End Function

Sub UpdateCurrentTestRecord(value As String)
    Select Case Trim(CurrentTestCaseProperty)
        Case "Tag:"
            CurrentTestCase.tag = value
        Case "Priority:"
            CurrentTestCase.priority = value
        Case "Environment:"
            CurrentTestCase.environment = value
        Case "Execution Type:"
            CurrentTestCase.execute_type = value
        Case "Test Type:"
            CurrentTestCase.test_type = value
        Case "Requirement:"
            CurrentTestCase.requirement = value
        Case "Reference:"
            CurrentTestCase.reference = value
        Case "Description:"
            CurrentTestCase.description = value
        Case "Precondition:"
            CurrentTestCase.precondition = value
        Case "Action:"
            CurrentTestCase.action = value
        Case "Result:"
            CurrentTestCase.result = value
        Case Else
            'MsgBox "Unknown property found " + CurrentTestCaseProperty + " with value " + value
    End Select
End Sub

Function IsValidTestRecord() As Boolean
    IsValidTestRecord = True
    With CurrentTestCase
        If .slogen = "" Then
            IsValidTestRecord = False
            Exit Function
        End If
        
        If .tag = "" Then
            IsValidTestRecord = False
            Exit Function
        End If
    End With
End Function

Sub ResetTestRecord()
    With CurrentTestCase
        .slogen = ""
        .tag = ""
        .priority = ""
        .environment = ""
        .execute_type = ""
        .test_type = ""
        .requirement = ""
        .reference = ""
        .description = ""
        .precondition = ""
        .action = ""
    End With
End Sub

Function SafeSql(str As String) As String
    SafeSql = Replace(str, "'", "''")
End Function

Sub SaveTestRecord()
On Error GoTo Err_Insert
    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    conn.Open "DSN=testcase;" & _
        "Database=testcase;" & _
        "Uid=postgres;" & _
        "Pwd=123456"
    Dim sql As String

    sql = "INSERT INTO cases (slogen,tag,priority,environment,exection_type,test_type,requirement,reference,description,precondition,action,result)" & vbCrLf & _
          "VALUES (?,?,?,?,?,?,?,?,?,?,?,?);"
    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter("slogen", adVarChar, adParamInput, 256, CurrentTestCase.slogen)
        .Parameters.Append .CreateParameter("tag", adVarChar, adParamInput, 64, CurrentTestCase.tag)
        .Parameters.Append .CreateParameter("priority", adChar, adParamInput, 8, CurrentTestCase.priority)
        .Parameters.Append .CreateParameter("environment", adChar, adParamInput, 8, CurrentTestCase.environment)
        .Parameters.Append .CreateParameter("exection_type", adChar, adParamInput, 8, CurrentTestCase.execute_type)
        .Parameters.Append .CreateParameter("test_type", adChar, adParamInput, 8, CurrentTestCase.test_type)
        .Parameters.Append .CreateParameter("requirement", adVarChar, adParamInput, 64, CurrentTestCase.requirement)
        .Parameters.Append .CreateParameter("reference", adVarChar, adParamInput, 64, CurrentTestCase.reference)
        .Parameters.Append .CreateParameter("description", adVarChar, adParamInput, 4096, CurrentTestCase.description)
        .Parameters.Append .CreateParameter("precondition", adVarChar, adParamInput, 2048, CurrentTestCase.precondition)
        .Parameters.Append .CreateParameter("action", adVarChar, adParamInput, 4096, CurrentTestCase.action)
        .Parameters.Append .CreateParameter("result", adVarChar, adParamInput, 2048, CurrentTestCase.result)
        .CommandText = sql
        .Execute
    End With
    conn.Close
    

Exit_Insert:
    Set cmd = Nothing
    Set rs = Nothing
    Set conn = Nothing
    Exit Sub
  
Err_Insert:
    MsgBox CurrentSection
    Resume Exit_Insert
    
End Sub

Function IsValidPropertyName(property As String) As Boolean
    Select Case Trim(property)
        Case "Tag:", "Priority:", "Environment:", "Execution Type:", "Test Type:", "Requirement:", "Reference:", "Description:", "Precondition:", "Action:", "Result:"
            IsValidPropertyName = True
        Case Else
            IsValidPropertyName = False
    End Select
End Function

Function GetTestCaseProperties() As Boolean
    Dim propertyName As String
    Dim propertyValue As String
    Dim selectText As String
    Dim unit As Integer
    Do While True
        unit = Selection.MoveDown
        If unit = 1 And IsSameTestCase Then
            propertyName = GetPropertyName
            If Not IsValidPropertyName(propertyName) Then
                GoTo nextCycle
            End If
            
            If propertyName = CurrentTestCaseProperty Or CurrentTestCaseProperty = "" Then
                If Not IsInTable Then
                    Selection.Expand wdLine
                    propertyValue = propertyValue + CleanString(Selection.Text)
                Else
                    propertyValue = propertyValue + TranslateTable
                End If
                Selection.Collapse
            Else
                UpdateCurrentTestRecord propertyValue
                Selection.Expand wdLine
                propertyValue = CleanString(Selection.Text)
            End If
            CurrentTestCaseProperty = propertyName
        Else
            UpdateCurrentTestRecord propertyValue
            If IsValidTestRecord Then
                SaveTestRecord
                ResetTestRecord
            End If
            propertyValue = Selection.Text
            Exit Do
        End If
nextCycle:
    Loop
    If unit = 0 Then
        GetTestCaseProperties = False
    Else
        GetTestCaseProperties = True
    End If
End Function

Function IsInTable() As Boolean
    ' Collapse the range to start so as to not have to deal with '
    ' multi-segment ranges. Then check to make sure cursor is '
    ' within a table. '
    Selection.Collapse Direction:=wdCollapseStart
    If Not Selection.Information(wdWithInTable) Then
        IsInTable = False
        Exit Function
    End If
    IsInTable = True
End Function

Function TranslateTable$()
    ' Process every row in the current table. '
    Dim row As Integer
    Dim col As Integer
    Dim totalRowNum As Integer
    Dim rng As Range
    Dim cellNum As Integer
    
    totalRowNum = Selection.Tables(1).Rows.Count
    TranslateTable = "<table>"
    For row = 1 To totalRowNum
        TranslateTable = TranslateTable + "<tr>"
        ' Get the range for the leftmost cell. '
        For col = 1 To Selection.Tables(1).Columns.Count
            Set rng = Selection.Tables(1).Rows(row).Cells(col).Range
            TranslateTable = TranslateTable + "<td>" + CleanString(rng.Text) + "</td>"
            cellNum = cellNum + 1
        Next
        TranslateTable = TranslateTable + "</tr>"
    Next
    TranslateTable = TranslateTable + "</table>"
    Do While IsInTable
        Selection.MoveDown
    Loop
    Selection.MoveUp
End Function

Function CleanString(StrIn As String) As String
    Dim iCh As Integer
    CleanString = Trim(StrIn)
    For iCh = 1 To Len(StrIn)
        If Asc(Mid(StrIn, iCh, 1)) < 32 Then
            'remove special character
            CleanString = Left(StrIn, iCh - 1) & CleanString(Mid(StrIn, iCh + 1))
            Exit Function
        End If
    Next iCh
End Function

Function ParseTestCase() As Boolean
    CurrentTestCaseProperty = ""
    CurrentSection = ""
    ResetTestRecord
    
    GetTestCaseSlogan
    ParseTestCase = GetTestCaseProperties
End Function

Sub FetchTestCases()
    CurrentTestCaseProperty = ""
    CurrentSection = ""
    ResetTestRecord
    
    Selection.GoTo What:=wdGoToHeading, Which:=wdGoToFirst
    Do While ParseTestCase
    Loop
    MsgBox "Parse finished"
End Sub
