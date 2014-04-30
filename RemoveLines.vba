Sub RemoveCompletes()
'
' RemoveCompletes Macro
'
' Greg Klotz, May 2012
'
' Removes all lines in the document ending in C

    Dim singleLine As Paragraph
    Dim lineText As String
    Dim lineEnd As String

    For Each singleLine In ActiveDocument.Paragraphs
        'grab text String from Paragraph object
        lineText = singleLine.Range.Text
        'grab right most 2 characters (char + newline/paragraph ^p)
        lineEnd = Right(lineText, 2)
        'grab left most 1 char, don't want the special char at end
        lineEnd = Left(lineEnd, 1)

        'delete lines ending in C
        If lineEnd = "C" Then singleLine.Range.Delete
            
        Next singleLine
End Sub
