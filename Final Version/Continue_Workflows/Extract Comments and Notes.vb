Sub ExtractNotesAndComments()
    Dim ExcelComment As Comment
    Dim CreateCommetsTab As Boolean
    Dim CommentSheet As Worksheet
    CreateCommetsTab = TRUE
    Dim CommentsRow As Long
    Dim ModernComment As CommentThreaded
    Dim myList      As ListObject
    Dim ListCols    As Long
    Dim dtToday     As String
    Dim noteTabName As String
    
    dtToday = Format(Date, "mm.dd")
    
    noteTabName = "Notes-" & dtToday
    On Error Resume Next
    Sheets(noteTabName).Delete
    
    Sheets.Add(Before:=Sheets(1)).Name = noteTabName
    
    Set CommentSheet = Worksheets(noteTabName)
    CommentSheet.Cells.Clear
    
    CommentsRow = 2
    Dim Filepath    As String
    Filepath = ActiveWorkbook.Path
    CommentSheet.Activate
    
    CommentSheet.Range("A1").Value = "Note Location"
    CommentSheet.Range("B1").Value = "Cell Value"
    CommentSheet.Range("C1").Value = "Author"
    CommentSheet.Range("D1").Value = "Note"
    With CommentSheet.Range("A1:D1")
        .Font.Bold = TRUE
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(250, 70, 22)
        .Columns.ColumnWidth = 30
    End With
    
    For Each ws In Worksheets
        For Each ExcelComment In ws.Comments
            CommentSheet.Hyperlinks.Add Range("A" & CommentsRow), Address:="", SubAddress:="'" & ws.Name & "'!" & ExcelComment.Parent.Address, TextToDisplay:="'" & ws.Name & "'!" & ExcelComment.Parent.Address
            CommentSheet.Range("B" & CommentsRow).Value = ExcelComment.Parent.Value
            CommentSheet.Range("C" & CommentsRow).Value = ExcelComment.Author
            CommentSheet.Range("D" & CommentsRow).Value = ExcelComment.Text
            CommentsRow = CommentsRow + 1
        Next ExcelComment
    Next ws
    
    If commentsRow = 2 Then
        CommentSheet.Range("A"&commentsrow).Value = "No Notes found in workbook"
    End If
    
    With CommentSheet
        .Range("A1").CurrentRegion.Select
        .ListObjects.Add (xlSrcRange)
    End With
    
    Set myList = CommentSheet.ListObjects(1)
    myList.TableStyle = "TableStyleLight8"
    ListCols = myList.DataBodyRange _
               .Columns.Count
    
    With myList.DataBodyRange
        .Cells.VerticalAlignment = xlTop
        .Columns.EntireColumn.ColumnWidth = 30
        .Cells.WrapText = TRUE
        .Columns.EntireColumn.AutoFit
        .Rows.EntireRow.AutoFit
    End With
    
    Dim commentTabName As String
    commentTabName = "Comments-" & dtToday
    
    On Error Resume Next
    Sheets(commentTabName).Delete
    
    Sheets.Add(Before:=Sheets(1)).Name = commentTabName
    
    Application.ScreenUpdating = FALSE
    
    Dim myCmt       As CommentThreaded
    Dim myRp        As CommentThreaded
    Dim curwks      As Worksheet
    Dim newwks      As Worksheet
    Dim i           As Long
    Dim iR          As Long
    Dim iRCol       As Long
    Dim cmtCount    As Long
    
    Set newwks = Worksheets(commentTabName)
    
    CommentsRow = 2
    
    newwks.Range("A1:F1").Value = _
                                  Array("Comment Location", "Cell Value", "Author", _
                                  "Date", "Replies", "Comment Text")
    With newwks.Range("A1:F1")
        .Font.Bold = TRUE
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(250, 70, 22)
        .Columns.ColumnWidth = 30
    End With
    
    For Each ws In Worksheets
        For Each ModernComment in ws.CommentsThreaded
            
            newwks.Hyperlinks.Add Range("A" & CommentsRow), Address:="", SubAddress:="'" & ws.Name & "'!" & ModernComment.Parent.Address, TextToDisplay:="'" & ws.Name & "'!" & ModernComment.Parent.Address
            newwks.Range("B" & CommentsRow).Value = ModernComment.Parent.Value
            newwks.Range("C" & CommentsRow).Value = ModernComment.Author.Name
            newwks.Range("D" & CommentsRow).Value = ModernComment.Date
            newwks.Range("E" & CommentsRow).Value = ModernComment.Replies.Count
            newwks.Range("F" & CommentsRow).Value = ModernComment.Text
            If ModernComment.Replies.Count >= 1 Then
                iR = 1
                iRCol = 7
                For Each r In ModernComment.Replies
                    newwks.Cells(1, iRCol).Value = "Reply " & iR
                    newwks.Cells(CommentsRow, iRCol).Value _
                                              = r.Author.Name _
                                            & vbCrLf _
                                            & r.Date _
                                            & vbCrLf _
                                            & r.text
                    iRCol = iRCol + 1
                    iR = iR + 1
                Next
            End If
            CommentsRow = CommentsRow + 1
        Next ModernComment
    Next ws
    
    If commentsRow = 2 Then
        newwks.Range("A"&commentsrow).Value = "No Comments found in workbook"
    End If
    
    With newwks
        .Range("A1").CurrentRegion.Select
        .ListObjects.Add (xlSrcRange)
    End With
    
    Set myList = newwks.ListObjects(1)
    myList.TableStyle = "TableStyleLight8"
    ListCols = myList.DataBodyRange _
               .Columns.Count
    
    With myList.DataBodyRange
        .Cells.VerticalAlignment = xlTop
        .Columns.EntireColumn.ColumnWidth = 30
        .Cells.WrapText = TRUE
        .Columns.EntireColumn.AutoFit
        .Rows.EntireRow.AutoFit
    End With
    
    Application.ScreenUpdating = TRUE
    
    Application.DisplayAlerts = FALSE
    Application.DisplayAlerts = TRUE
    
    ActiveWorkbook.Save
    
End Sub