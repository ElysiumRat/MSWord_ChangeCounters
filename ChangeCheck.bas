Attribute VB_Name = "ChangeCheck"
Sub ChangeCheck()

Dim given, currentauthor As String
Dim countChg, countComm, totalChg, totalComm, counter, wordcount As Long

given = InputBox("Please type the username of the editor you wish to evaluate.")

wordcount = ActiveDocument.Range.ComputeStatistics(wdStatisticWords)

totalChg = 0
countChg = ActiveDocument.Range.Revisions.Count
countComm = ActiveDocument.Range.comments.Count

For counter = 1 To countChg
    currentauthor = ActiveDocument.Range.Revisions(counter).Author
    If currentauthor = given Then
        totalChg = totalChg + 1
    End If
Next counter

For counter = 1 To countComm
    currentauthor = ActiveDocument.Range.comments(counter).Author
    If currentauthor = given Then
        totalComm = totalComm + 1
    End If
Next counter

MsgBox "Editor name: " & given & vbCrLf & "Changes by this editor: " & totalChg & vbCrLf & "Comments by this editor: " & totalComm

End Sub