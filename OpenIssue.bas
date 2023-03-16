' Calls OpenUrl Utils -> ShellExecute

Public Sub OpenLinkIssue(indCol As Integer)
' indCol: Column index of Mapping table triggered

Dim tblMapping As ListObject
Dim tblSelected As ListObject
Dim rng As Range
Dim strID As String
Dim CurrentCell As Range

Set tblSelected = ActiveCell.ListObject
' Get ID Table
Set tblMapping = ThisWorkbook.Sheets("ID Mapping").ListObjects("IDTable")
Set rng = tblMapping.Range

If tblSelected Is Nothing Then
    
    With rng.Columns(indCol)
        Set Found = .Find(what:="*", after:=.Cells(1, 1), LookIn:=xlFormulas) ' Search for non empty/ skip header
        If Found Is Nothing Then
            MsgBox "Hotkey defined without links!"
            Exit Sub
        End If
    End With
    strPreLink = Found.Value
    Set CurrentCell = ActiveCell
    GoTo OpenIssues
    
End If

Dim strCurrentColName As String
' http://stackoverflow.com/questions/7643652/finding-active-cells-column-header-name-using-vba-in-excel
strCurrentColName = Cells(tblSelected.Range.Row, ActiveCell.Column).Value
' Partial match e.g. "JIRA-ID (Application)" matches "JIRA-ID"
' => Ignore after first white space
strCurrentColName = SplitRe(strCurrentColName, "[\s]")(0)

'Debug.Print "strCurrentColName: " & strCurrentColName

With rng.Columns(1)

    Set Found = .Find(what:=strCurrentColName)
    If Found Is Nothing Then
        ' Look for default ID in header = first filled one
        With rng.Columns(indCol)
            Set Found = .Find(what:="*", after:=.Cells(1, 1), LookIn:=xlFormulas) ' Search for non empty/ skip header
            If Found Is Nothing Then
                MsgBox "Hotkey defined without links!"
                Exit Sub
            End If
        End With
        indRow = Found.Row - tblMapping.HeaderRowRange.Row + 1 ' relative index
        strDefID = rng.Cells(indRow, 1) ' First Colum
        strPreLink = Found.Value
              
        With tblSelected.Range.Rows(1)
            Set Found = .Find(what:=strDefID)
            If Found Is Nothing Then
                MsgBox "'" & strCurrentColName & "' not found in Mapping Table and '" _
                & strDefID & "' (hotkey default) not found in selected Table header!"
                Exit Sub
            End If
            'col = Found.Column - tblSelected.HeaderRowRange.Column + 1
            Set CurrentCell = tblSelected.Range.Cells(ActiveCell.Row - tblSelected.HeaderRowRange.Row + 1, _
            Found.Column - tblSelected.HeaderRowRange.Column + 1)
            
        End With
    Else
        Set CurrentCell = ActiveCell
        indRow = Found.Row - tblMapping.HeaderRowRange.Row + 1 ' relative index
        strPreLink = tblMapping.Range.Cells(indRow, indCol).Value
    End If
End With

OpenIssues:

' Check if empty
If IsEmpty(strPreLink) Then
    MsgBox "Link is not defined for Hotkey and ID."
    Exit Sub
End If

strID = Trim(CurrentCell.Value)

' Loop if multiple IDs entered
' Calls Utils/SplitRe
Dim strArrayID() As String
strArrayID = SplitRe(strID, "[\s]+")

For I = 0 To UBound(strArrayID)
'    If I <> 0 Then
'        Set WshShell = WScript.CreateObject("WScript.Shell")
'        WshShell.Sleep 2000 'milliseconds
'    End If

' Check if contains numbers. If not skip.
    If HasNumber(strArrayID(I)) Then
        strLink = strPreLink & strArrayID(I)
        ' MsgBox strLink
        'OpenUrl (strLink)
        ThisWorkbook.FollowHyperlink Address:=strLink
    Else
        Debug.Print "Skip open issue '" & strArrayID(I) & "' because no number in the string."
    End If
Next I
    
End Sub


Sub LoadOpenIssueHotkeys()
Dim tblMapping As ListObject
Dim strKey As String

Set tblMapping = ThisWorkbook.Sheets("ID Mapping").ListObjects("IDTable")
' Loop on list of OnKey
Set tblHeader = tblMapping.HeaderRowRange


For I = 2 To tblHeader.Columns.Count
    strKey = tblHeader(1, I).Value
   Application.OnKey strKey, "'OpenLinkIssue " & I & "'"
Next I
End Sub

Sub OpenIssueCloseAddInWorkbook()
' TODO Check if xlsa file
ThisWorkbook.IsAddin = False
' TODO Open Sheet
End Sub

Sub OpenIssueLoadHotkeys()
Dim tblMapping As ListObject
Dim strKey As String

Set tblMapping = ThisWorkbook.Sheets("ID Mapping").ListObjects("IDTable")
' Loop on list of OnKey
Set tblHeader = tblMapping.HeaderRowRange


For I = 2 To tblHeader.Columns.Count
    strKey = tblHeader(1, I).Value
   Application.OnKey strKey, "'OpenLinkIssue " & I & "'"
Next I
End Sub


Function HasNumber(strData As String) As Boolean
    For iCnt = 1 To Len(strData)
        If IsNumeric(Mid(strData, iCnt, 1)) Then
            HasNumber = True
            Exit Function
        End If
    Next

    HasNumber = False
End Function

' Not Used

Function TableHeader(cl As Range) As Variant
    Dim lst As ListObject
    Dim strHeading As String

    Set lst = cl.ListObject

    If Not lst Is Nothing Then
        TableHeader = lst.HeaderRowRange.Cells(1, cl.Column - lst.Range.Column + 1).Value
    Else
        TableHeader = ""
    End If
End Function
