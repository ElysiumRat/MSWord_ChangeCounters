Attribute VB_Name = "SelectionCheck"
Sub SelectionCheck()

Dim selrange As Range
Dim given, currentauthor As String
Dim countChg, countComm, counter, totalChg, totalComm, selstart, selend, wordcount As Long

given = InputBox("Please type the username of the editor you wish to evaluate.")

Application.ScreenUpdating = False

Set selrange = ActiveDocument.Range
selstart = Selection.Start
selend = Selection.End
selrange.SetRange Start:=selstart, End:=selend
selrange.Select

wordcount = Selection.Range.ComputeStatistics(wdStatisticWords)

totalChg = 0
totalComm = 0
countChg = selrange.Revisions.Count
countComm = selrange.comments.Count

For counter = 1 To countChg
    currentauthor = selrange.Revisions(counter).Author
    If currentauthor = given Then
        totalChg = totalChg + 1
    End If
Next counter

For counter = 1 To countComm
    currentauthor = selrange.comments(counter).Author
    If currentauthor = given Then
        totalComm = totalComm + 1
    End If
Next counter

MsgBox "Editor: " & given & vbCrLf & "Changes by this editor in selection: " & totalChg & vbCrLf & "Comments by this editor in selection: " & totalComm & vbCrLf & "Word count of selection: " & wordcount
End Sub






