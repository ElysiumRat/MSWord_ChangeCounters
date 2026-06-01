# AuthorCheck
A VBA macro for Microsoft Word to count the numbers of changes made by specific authors.

`AuthorCheck` scans all comments and revisions in the active Word document, groups them by author, and outputs a summary showing:

- Number of revisions per author; 
- Number of comments per author.

It uses a `Scripting.Dictionary` to store authors and counters, avoiding the complexity and fragility of manually managed arrays.

I previously attempted to make [a Word VBA macro for counting changes in a document by author](https://github.com/ElysiumRat/MSWord_ChangeCounters). My prior attempts failed for reasons related to Word's backend, and further research has given me a working solution, which I provide separately here for sake of clarity.

## How AuthorCheck Works

### 1. Creates a dictionary to store authors
```vb
Dim dictAuthors As Object
Set dictAuthors = CreateObject("Scripting.Dictionary")
```

Each dictionary entry uses the author name as the key and a 2‑element array as the value:

```
(0) = number of revisions  
(1) = number of comments
```

### 2. First pass: collect all authors
```vb
For Each c In ActiveDocument.Comments
    author = c.Author
    If Not dictAuthors.Exists(author) Then
        dictAuthors.Add author, Array(0, 0)
    End If
Next c
```

The same logic is applied to revisions. This ensures each author is added exactly once.

### 3. Second pass: count comments and revisions
```vb
arrC = dictAuthors(author)
arrC(1) = arrC(1) + 1
dictAuthors(author) = arrC
```

The macro updates the counters cleanly without resizing arrays or searching for matching authors.

### 4. Output the results
```vb
For Each key In dictAuthors.Keys
    output = output & "Editor: " & key & vbCrLf & _
             "Changes: " & dictAuthors(key)(0) & vbCrLf & _
             "Comments: " & dictAuthors(key)(1) & vbCrLf & vbCrLf
Next key
```

## How AuthorCheck Differs From the Older Macros

### 1. No manual array resizing
Older macros used patterns like:

```vb
ReDim Preserve AuthorList(ArrayPos)
```

This is slow, fragile, and easy to break. `AuthorCheck` avoids arrays entirely by using a dictionary.

### 2. No loop‑counter corruption
`AuthorArrayCheck` contained logic like:

```vb
CountArray = ArrayPos
```

This causes unpredictable behavior. `AuthorCheck` uses `For Each`, which avoids index manipulation entirely.

### 3. No duplicate‑detection logic
Older macros manually searched arrays to check if an author already existed.

`AuthorCheck` replaces all of that with:

```vb
If Not dictAuthors.Exists(author) Then ...
```

This is faster and simpler.

### 4. No reliance on user‑typed author names
`ChangeCheck` and `SelectionCheck` required the user to type an author name exactly as Word stores it.

`AuthorCheck` automatically discovers all authors, eliminating user error.

### 5. No selection‑based revision counting
`SelectionCheck` depended on `Selection.Range.Revisions`, which is unreliable.

`AuthorCheck` always uses the full document, where Word’s revision model is stable.

### 6. No duplicated logic
Older macros repeated the same loops multiple times.

`AuthorCheck` uses a clean two‑pass structure:

1. Collect authors;
2. Count items.
