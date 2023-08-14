Attribute VB_Name = "AuthorArrayCheck"
Sub AuthorArrayCheck()

Dim AuthorList() As String
Dim tempauthor, currentauthor, comp As String
Dim CountComm, countChg, CountArray, ArrayPos, counter As Integer
Dim NumChg(), NumComm() As Integer
Dim Found As Boolean
Dim temp As String

CountComm = ActiveDocument.Range.comments.Count
countChg = ActiveDocument.Range.Revisions.Count

ReDim Preserve AuthorList(1)
currentauthor = ActiveDocument.Range.comments(1).Author
AuthorList(0) = currentauthor
ArrayPos = 0

For counter = 1 To CountComm
    Found = False
    currentauthor = ActiveDocument.Range.comments(counter).Author
    For CountArray = 0 To ArrayPos
        tempauthor = AuthorList(CountArray)
        If tempauthor = currentauthor Then
            Found = True
        End If
    Next CountArray
    If Found = False Then
        ArrayPos = ArrayPos + 1
        ReDim Preserve AuthorList(ArrayPos)
        AuthorList(ArrayPos) = currentauthor
    End If
Next counter

' Getting a list of the names from changes and comments

For counter = 1 To countChg
    Found = False
    currentauthor = ActiveDocument.Range.Revisions(counter).Author
    For CountArray = 0 To ArrayPos
        tempauthor = AuthorList(CountArray)
        If tempauthor = currentauthor Then
            Found = True
            CountArray = ArrayPos
        End If
    Next CountArray
    If Found = False Then
        ArrayPos = ArrayPos + 1
        ReDim Preserve AuthorList(ArrayPos)
        AuthorList(ArrayPos) = currentauthor
    End If
Next counter

For counter = 1 To CountComm
    Found = False
    currentauthor = ActiveDocument.Range.comments(counter).Author
    For CountArray = 0 To ArrayPos
        tempauthor = AuthorList(CountArray)
        If tempauthor = currentauthor Then
            Found = True
            CountArray = ArrayPos
        End If
    Next CountArray
    If Found = False Then
        ArrayPos = ArrayPos + 1
        ReDim Preserve AuthorList(ArrayPos)
        AuthorList(ArrayPos) = currentauthor
    End If
Next counter

ReDim NumChg(ArrayPos)
ReDim NumComm(ArrayPos)

' Giving default values for the arrays, for display purposes
For counter = 0 To ArrayPos
    NumChg(counter) = 0
    NumComm(counter) = 0
Next counter

' Counting the numbers of comments and changes by author

For counter = 1 To countChg
    currentauthor = ActiveDocument.Range.Revisions(counter).Author
    For CountArray = 0 To ArrayPos
        tempauthor = AuthorList(CountArray)
        If tempauthor = currentauthor Then
            NumChg(CountArray) = NumChg(CountArray) + 1
        End If
    Next CountArray
Next counter

For counter = 1 To CountComm
    currentauthor = ActiveDocument.Range.comments(counter).Author
    For CountArray = 0 To ArrayPos
        tempauthor = AuthorList(CountArray)
        If tempauthor = currentauthor Then
            NumComm(CountArray) = NumComm(CountArray) + 1
        End If
    Next CountArray
Next counter

' Constructing the dialogue to output the information

temp = "Editor: " & AuthorList(0) & vbCrLf & "Changes: " & NumChg(0) & vbCrLf & "Comments: " & NumComm(0)

For counter = 1 To ArrayPos
    temp = temp & vbCrLf & vbCrLf & "Editor: " & AuthorList(counter) & vbCrLf & "Changes: " & NumChg(counter) & vbCrLf & "Comments: " & NumComm(counter)
Next counter

MsgBox temp

End Sub
