Attribute VB_Name = "AuthorCheck"
Sub AuthorCheck()

    Dim dictAuthors As Object
    Set dictAuthors = CreateObject("Scripting.Dictionary")

    Dim c As Comment
    Dim r As Revision
    Dim author As String

    ' First pass: collect all authors and initialize counters
    For Each c In ActiveDocument.Comments
        author = c.author
        If Not dictAuthors.Exists(author) Then
            dictAuthors.Add author, Array(0, 0) ' (NumChg, NumComm)
        End If
    Next c

    For Each r In ActiveDocument.Revisions
        author = r.author
        If Not dictAuthors.Exists(author) Then
            dictAuthors.Add author, Array(0, 0)
        End If
    Next r

    ' Second pass: count comments
    For Each c In ActiveDocument.Comments
        author = c.author
        Dim arrC As Variant
        arrC = dictAuthors(author)
        arrC(1) = arrC(1) + 1
        dictAuthors(author) = arrC
    Next c

    ' Second pass: count revisions
    For Each r In ActiveDocument.Revisions
        author = r.author
        Dim arrR As Variant
        arrR = dictAuthors(author)
        arrR(0) = arrR(0) + 1
        dictAuthors(author) = arrR
    Next r

    ' Build output
    Dim output As String
    Dim key As Variant

    For Each key In dictAuthors.Keys
        output = output & "Editor: " & key & vbCrLf & _
                 "Changes: " & dictAuthors(key)(0) & vbCrLf & _
                 "Comments: " & dictAuthors(key)(1) & vbCrLf & vbCrLf
    Next key

    MsgBox output

End Sub


