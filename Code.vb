Private Sub CommandButton1_Click()
    Dim SheetList() As Variant
    For I = 1 To ActiveWorkbook.Sheets.Count
        ReDim Preserve SheetList(0 To I) As Variant 'Resize list
        SheetList(I) = ActiveWorkbook.Sheets(I).Name
    Next I
    
    ComboBox1.List = SheetList
End Sub

Private Sub CommandButton2_Click()
    GenerateSchedule
End Sub

Public Sub GenerateSchedule()
'=== Setup layout ===
'Set height
Rows(1).RowHeight = Cells(9, "P")
Rows("2:42").RowHeight = Cells(9, "T") 'Including source text row

'Set width
Columns(1).ColumnWidth = Cells(9, "R")
Columns("B:K").ColumnWidth = Cells(9, "V")

'Set header grid
Cells(3, "P").Copy
Range("A1:K1").PasteSpecial xlPasteAll

'Set hour grid
For I = 2 To 38 Step 4
    'Merge cells
    Range(Cells(I, 1), Cells(I + 3, 1)).Merge
Next I

For I = 2 To 38 Step 8
    Range("R27:R34").Copy
    Range(Cells(I, 1), Cells(I + 7, 1)).PasteSpecial xlPasteAll
Next I

'Set lecture grid
For I = 2 To 41 Step 4
    For J = 2 To 11 Step 2
        Dim GroupRange As Range: Set GroupRange = Range(Cells(I, J), Cells(I + 3, J + 1))
        Cells(27, "P").Copy
        GroupRange.PasteSpecial xlPasteAll
        GroupRange.Merge
    Next J
Next I


'=== Setup header and hour text ===
'Set hour title and center the cell
Cells(1, "A").Value = Cells(14, "R").Value
Cells(1, "A").HorizontalAlignment = xlCenter
Cells(1, "A").VerticalAlignment = xlCenter
CopyFont Cells(1, "A").Characters(1, Len(Cells(1, "A").Value)), Cells(6, "P")

'Set header cells text
For I = 2 To 10 Step 2
    'Setup variables
    Dim HeaderRange As Range, HeaderText As String
    Set HeaderRange = Range(Cells(1, I), Cells(1, I + 1))
    HeaderText = Cells(14 + I \ 2, "R").Value
    
    'Set day text
    Cells(1, I).Value = HeaderText
    
    'Merge header cells
    HeaderRange.Merge
    
    'Center header text
    HeaderRange.HorizontalAlignment = xlCenter
    HeaderRange.VerticalAlignment = xlCenter
    
    'Style header
    CopyFont HeaderRange.Characters(1, Len(HeaderText)), Cells(6, "P")
Next I

'Set hour cells text
For RowIndex = 2 To 41 Step 4
    'Setup variables
    Dim HourRange As Range, HourText As String
    Set HourRange = Range(Cells(RowIndex, 1), Cells(RowIndex + 3, 1))
    HourText = Cells((RowIndex - 2) / 4 + 14, "X")

    'Set time peroid cell content
    Cells(RowIndex, 1).Value = HourText
    
    'Set font style
    CopyFont Cells(RowIndex, 1).Characters(1, Len(HourText)), Cells(6, "R")
    
    'Center and merge column content
    HourRange.Merge
    HourRange.HorizontalAlignment = xlCenter
    HourRange.VerticalAlignment = xlCenter
Next RowIndex

'=== Setup data ===
'Populate lecture cells
ParseDataTable (Sheets(ComboBox1.Value).ListObjects("Table1"))

'=== Setup main border ===
BorderParts = Array(xlEdgeBottom, xlEdgeLeft, xlEdgeRight, xlEdgeTop)
For Each BorderPart In BorderParts
    Range("A1:K41").Borders(BorderPart).Color = Cells(27, "T").Borders(BorderPart).Color
    Range("A1:K41").Borders(BorderPart).ColorIndex = Cells(27, "T").Borders(BorderPart).ColorIndex
    Range("A1:K41").Borders(BorderPart).LineStyle = Cells(27, "T").Borders(BorderPart).LineStyle
    Range("A1:K41").Borders(BorderPart).TintAndShade = Cells(27, "T").Borders(BorderPart).TintAndShade
    Range("A1:K41").Borders(BorderPart).Weight = Cells(27, "T").Borders(BorderPart).Weight
Next BorderPart

'=== Setup version text ===
Dim SourceText As String: SourceText = Cells(12, "P").Value
Cells(42, 1).Value = SourceText
Range(Cells(42, 1), Cells(42, "E")).Merge
CopyFont Cells(42, 1).Characters(1, Len(SourceText)), Cells(9, "X")

'=== Setup original schedule and syllabus hyperlinks ===
Worksheets(1).Hyperlinks.Add Anchor:=Range("F42:G42"), _
                             Address:=Cells(43, "P").Value, _
                             TextToDisplay:=Cells(30, "V").Value

Worksheets(1).Hyperlinks.Add Anchor:=Range("H42:I42"), _
                             Address:=Cells(37, "P").Value, _
                             TextToDisplay:=Cells(27, "V").Value

Worksheets(1).Hyperlinks.Add Anchor:=Range("J42:K42"), _
                             Address:=Cells(40, "P").Value, _
                             TextToDisplay:=Cells(27, "x").Value
               
Range("A43:K42").VerticalAlignment = xlCenter

CopyFont Cells(42, "F").Characters(1, Len(SourceText)), Cells(9, "X")
CopyFont Cells(42, "H").Characters(1, Len(SourceText)), Cells(9, "X")
CopyFont Cells(42, "J").Characters(1, Len(SourceText)), Cells(9, "X")

Range(Cells(42, "F"), Cells(42, "G")).Merge
Range(Cells(42, "H"), Cells(42, "I")).Merge
Range(Cells(42, "J"), Cells(42, "K")).Merge


End Sub

Function ParseDataTable(ByVal Table As ListObject)
    For Each Item In Table.ListRows
        Dim ScheduleRange As Range, LectureTypeInnerText As String
        Set ScheduleRange = FindRangeByTime(Item.Range(1), Item.Range(3), Item.Range(4), Item.Range(2))
        ScheduleRange.UnMerge
        
        'Copy style
        Select Case Item.Range(6)
        Case "L"
            LectureTypeInnerText = Cells(21, "R").Value
            Range("R3").Copy
        Case "C"
            LectureTypeInnerText = Cells(22, "R").Value
            Range("T3").Copy
        Case "P"
            LectureTypeInnerText = Cells(23, "R").Value
            Range("V3").Copy
        Case "W"
            LectureTypeInnerText = Cells(24, "R").Value
            Range("X3").Copy
        End Select
        
        'Paste into the lecture range
        ScheduleRange.PasteSpecial xlPasteAll 'xlPasteFormats xlPasteAllUsingSourceTheme
        
        'Setup text variables
        Dim TitleText, LecturerText, PlaceText, InfoText, ResultText
        TitleText = Item.Range(5) & vbCrLf
        LectureTypeText = ChrW(&H2014) & " " & LectureTypeInnerText & " " & ChrW(&H2014) & vbCrLf
        LecturerText = Item.Range(7)
        PlaceText = Item.Range(8)
        InfoText = " " & ChrW(&H2014) & " " & Item.Range(9)
        
        ResultText = TitleText & LectureTypeText & LecturerText & vbCrLf & PlaceText
                             
        If Item.Range(9) <> "" Then
            ResultText = ResultText & InfoText
        End If
        
        'Set and center the value; merge cells
        ScheduleRange.Cells(1, 1).Value = ResultText
        ScheduleRange.Cells(1, 1).HorizontalAlignment = xlCenter
        ScheduleRange.Cells(1, 1).VerticalAlignment = xlCenter
        ScheduleRange.Merge
        
        'TODO ================ Fix week 1 & 2
        
        Dim Len1, Len2, Len3
        Len1 = Len(TitleText)
        Len2 = Len(TitleText) + Len(LectureTypeText)
        Len3 = Len(TitleText) + Len(LectureTypeText) + Len(LecturerText)
        Len4 = Len(TitleText) + Len(LectureTypeText) + Len(LecturerText) + Len(PlaceText)
        Len5 = Len(TitleText) + Len(LectureTypeText) + Len(LecturerText) + Len(PlaceText) + Len(InfoText)
        
        'Copy fonts
        CopyFont ScheduleRange.Cells(1, 1).Characters(1, Len1), Cells(6, "T")
        CopyFont ScheduleRange.Cells(1, 1).Characters(Len1 + 1, Len2), Cells(6, "V")
        CopyFont ScheduleRange.Cells(1, 1).Characters(Len2 + 1, Len3), Cells(6, "V")
        CopyFont ScheduleRange.Cells(1, 1).Characters(Len3 + 1, Len4), Cells(6, "X")
        
        If Item.Range(9) <> "" Then
            CopyFont ScheduleRange.Cells(1, 1).Characters(Len4 + 1, Len5), Cells(6, "X")
        End If
        
        ScheduleRange.Dirty
    Next Item
End Function

Function FindRangeByTime(ByVal Day As Integer, _
                         ByVal StartTime As Date, _
                         ByVal EndTime As Date, _
                         ByVal Week As Integer) As Excel.Range
                         
If Day < 0 Or Day > 5 Then
    Err.Raise Number:=vbObjectError + 1, _
              Description:="Invalid Day"
              
ElseIf (Not Week = 1) And (Not Week = 2) And (Not Week = 12) Then
    Err.Raise Number:=vbObjectError + 4, _
              Description:="Invalid Week"
End If

Dim StartColumnIndex As Integer, EndColumnIndex As Integer
Select Case Week
Case 1
    StartColumnIndex = Day * 2
    EndColumnIndex = StartColumnIndex
Case 2
    StartColumnIndex = Day * 2 + 1
    EndColumnIndex = StartColumnIndex
Case 12
    StartColumnIndex = Day * 2
    EndColumnIndex = Day * 2 + 1
End Select

Dim StartRowIndex As Integer, EndRowIndex As Integer
StartRowIndex = 2 + (Hour(StartTime) - 8) * 4
EndRowIndex = 2 + (Hour(StartTime) - 8) * 4 + (DateDiff("n", StartTime, EndTime) / 15) - 1

Set FindRangeByTime = Range(Cells(StartRowIndex, StartColumnIndex), Cells(EndRowIndex, EndColumnIndex))
End Function

Function IsDivisible(X As Integer, D As Integer) As Boolean
    IsDivisible = ((X Mod D) = 0)
End Function

Sub CopyFont(Target As Characters, Source As Range)
    Target.Font.Name = Source.Font.Name
    Target.Font.Size = Source.Font.Size
    Target.Font.Bold = Source.Font.Bold
    Target.Font.Italic = Source.Font.Italic
    Target.Font.Color = Source.Font.Color
End Sub

Sub CopyBorders(Target, Source, ShouldCopyInnerBorders As Boolean)

End Sub



