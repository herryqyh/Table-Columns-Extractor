Attribute VB_Name = "NewMacros"

Public mystr, mystrs() As String
Public history As String
Dim seq()


Sub 提取表格()
Attribute 提取表格.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.宏1"
'
    Dim tCount
    Dim tempstr() As String
    Dim i, j As Integer
    
    'MsgBox ActiveDocument.Name
    UserForm1.Show
    'str = InputBox("输入需要提取的列", , str)
    
    If mystr = "" Then
        MsgBox "输入为空"
        Exit Sub
    End If
    
    tempstr = Split(mystr, "(")
    mystrs = Split(tempstr(0), " ")
    
    
    ReDim Preserve seq(UBound(mystrs) - 1)
    For i = 0 To UBound(mystrs) - 1
        seq(i) = Val(mystrs(i))
    Next
    
    tCount = ActiveDocument.Tables.Count
    
    Application.ScreenUpdating = False
    Set MyRange = ActiveDocument.Content
    MyRange.Collapse Direction:=wdCollapseEnd
    
    Selection.EndKey Unit:=wdStory
    Selection.InsertBreak Type:=wdPageBreak
    For i = 1 To tCount
        Selection.EndKey Unit:=wdStory
        Selection.TypeParagraph
        ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=ActiveDocument.Tables(1).Rows.Count, NumColumns:=UBound(seq) + 1
    Next
    
    For j = 1 To tCount
        For i = 0 To UBound(seq)
            ActiveDocument.Tables(j).Columns(seq(i)).Select
            Selection.Copy
            ActiveDocument.Tables(j + tCount).Columns(i + 1).Select
            
            Selection.PasteAndFormat Type:=wdFormatOriginalFormatting
            ActiveDocument.Tables(j + tCount).Columns(i + 1).Width = InchesToPoints(0.8)
            
            ActiveDocument.Tables(j + tCount).Columns(i + 2).Select
            Selection.Columns.Delete
            
            ActiveDocument.Tables(j + tCount).Columns.Select
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Next
    Next
End Sub

Sub 提取表格batch()
    Dim fd As FileDialog, vrtSelectedItem As Variant, iFile As Document
    Dim tempstr() As String
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = True
        .InitialFileName = ActiveDocument.Path
        .Filters.Add "Documents", "*.doc; *.docx; *.rtf", 1
        .FilterIndex = 2
        If .Show <> -1 Then
            MsgBox "您没有选择任何文档！", vbCritical
            Exit Sub
        Else
            UserForm1.Show

            If mystr = "" Then
                MsgBox "输入为空"
                Exit Sub
            End If
            
            tempstr = Split(mystr, "(")
            mystrs = Split(tempstr(0), " ")
            
            ReDim Preserve seq(UBound(mystrs) - 1)
            For i = 0 To UBound(mystrs) - 1
                seq(i) = Val(mystrs(i))
            Next
            
            
            For Each vrtSelectedItem In .SelectedItems
                Set iFile = Documents.Open(vrtSelectedItem)
                iFile.Activate
                Call 提取表格_sub
                Application.DisplayAlerts = False
                iFile.Close True
                Application.DisplayAlerts = False
'                MsgBox "Selected item's path: " & vrtSelectedItem
            Next vrtSelectedItem
        End If
    End With
    Set iFile = Nothing
    Set fd = Nothing
    MsgBox "ok"
End Sub

Sub 提取表格_sub()
'
'
    'Dim seq = New Integer() {1, 2, 3, 8, 4}
    'Dim seq()
    Dim tCount
    'Dim tempstr() As String
    'Dim str As String
    Dim i, j As Integer
    
    'MsgBox ActiveDocument.Name
    '''
    'UserForm1.Show
    'str = InputBox("输入需要提取的列", , str)
    
    
    'MsgBox mystr
    'If mystr = "" Then
     '   MsgBox "输入为空"
     '   Exit Sub
    'End If
    '''
    'tempstr = Split(mystr, "(")
    'MsgBox tempstr(0)
    'mystrs = Split(tempstr(0), " ")
    'MsgBox mystr
    
    '''
    'ReDim Preserve seq(UBound(mystrs) - 1)
    For i = 0 To UBound(mystrs) - 1
        seq(i) = Val(mystrs(i))
    Next
    '''
    tCount = ActiveDocument.Tables.Count
    
    Application.ScreenUpdating = False
    Set MyRange = ActiveDocument.Content
    MyRange.Collapse Direction:=wdCollapseEnd
    
    Selection.EndKey Unit:=wdStory
    Selection.InsertBreak Type:=wdPageBreak
    For i = 1 To tCount
        Selection.EndKey Unit:=wdStory
        Selection.TypeParagraph
        ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=ActiveDocument.Tables(1).Rows.Count, NumColumns:=UBound(seq) + 1
    Next
    
    For j = 1 To tCount
        For i = 0 To UBound(seq)
            ActiveDocument.Tables(j).Columns(seq(i)).Select
            Selection.Copy
            ActiveDocument.Tables(j + tCount).Columns(i + 1).Select
            
            Selection.PasteAndFormat Type:=wdFormatOriginalFormatting
            ActiveDocument.Tables(j + tCount).Columns(i + 1).Width = InchesToPoints(0.8)
            
            ActiveDocument.Tables(j + tCount).Columns(i + 2).Select
            Selection.Columns.Delete
            
            ActiveDocument.Tables(j + tCount).Columns.Select
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Next
    Next
End Sub

