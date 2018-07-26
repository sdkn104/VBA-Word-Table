Attribute VB_Name = "TableUtil"
'
' Utilities for Word Table
'
'

'Sub TEST_SAMPLE()
'
'  Dim cellMap() As TableCell
'  Dim cm As TableCell
'  Dim cl As cell
'  Dim r As Long, c As Long
'
'  cellMap = CreateCellMap(ActiveDocument.Tables(1), r, c)
'  Debug.Print r & c
'
'  Debug.Print "===="
'  For r = 1 To UBound(cellMap, 1)
'    For c = 1 To UBound(cellMap, 2)
'      'If cellMap(r, c).IsMerged Then
'        cm = cellMap(r, c)
'        Debug.Print "(" & cm.Top & cm.Left & ")-(" & cm.Bottom & cm.Right & ") " & cm.IsMerged & ":" & GetTableCellText(cm.obj)
'      'End If
'    Next
'  Next
'  Debug.Print "----"
'
'  For Each cl In ActiveDocument.Tables(1).Range.Cells
'    Debug.Print cl.RowIndex & cl.ColumnIndex
'    cm = GetTableCell(cl, cellMap)
'    Debug.Print "(" & cm.Top & cm.Left & ")-(" & cm.Bottom & cm.Right & ") " & cm.IsMerged & ":" & GetTableCellText(cm.obj)
'  Next
'End Sub



' Type TableCell -- an element of the cell map
Public Type TableCell
    Left As Long  'Leftmost column index
    Right As Long  'Rightmost column index
    Top As Long  'Top row index
    Bottom As Long 'Bottom row index
    obj As Word.cell 'Word.Cell Object
    IsMerged As Boolean 'True if the cell is merged with another cell and not the top-left cell in the merged cells
End Type


'Create Cell Map
Public Function CreateCellMap(tbl As Table, Optional ByRef RowCount As Long = -1, Optional ByRef ColumnCount As Long = -1) As TableCell()
    Dim d As Object
    Dim colSpan As Long, colIdx As Long
    Dim i As Long, j As Long, k As Long, l As Long
    Dim col0 As Long, row0 As Long
    Dim RowCount_ As Long, ColumnCount_ As Long
    Dim cellMap_() As TableCell
    Dim cellMap2_() As TableCell
    
    'read XML and create CellMap
    Set d = CreateObject("MSXML2.DOMDocument")
    If Not d.LoadXML(tbl.Range.XML) Then Err.Raise 9999, "", "program error 2"
    With d.SelectNodes("/w:wordDocument/w:body/wx:sect/w:tbl/w:tr") 'exclude nested (inner) tables
      'init cell map array
      RowCount_ = .Length
      ColumnCount_ = tbl.Columns.Count
      ReDim cellMap_(1 To RowCount_, 1 To ColumnCount_)
      'create cell map
      For i = 1 To .Length
        colIdx = 1
        With .Item(i - 1).SelectNodes("./w:tc")
          For j = 1 To .Length
            'get colSpan
            colSpan = 1
            If .Item(j - 1).SelectNodes("w:tcPr/w:gridSpan").Length > 0 Then ' horizontally merged
              colSpan = CLng(.Item(j - 1).SelectNodes("w:tcPr/w:gridSpan").Item(0).Attributes(0).Text)
            End If
            'get vMerge
            Set vm = .Item(j - 1).SelectNodes("w:tcPr/w:vmerge | w:tcPr/w:vMerge")
            vmerge = 0
            If vm.Length > 0 Then
              vmerge = 1
              If vm(0).Attributes.Length > 0 Then
                If vm(0).Attributes(0).Value = "restart" Then vmerge = 2 'top of merged cells
              End If
            End If
            'update colum size
            If colIdx > ColumnCount_ Then
              ColumnCount_ = colIdx
              ReDim Preserve cellMap_(RowCount_, ColumnCount_)
            End If
            'set cell
            If vmerge = 1 Then ' merged with the upper cell
              colSpan = cellMap_(i - 1, colIdx).Right - cellMap_(i - 1, colIdx).Left + 1
              row0 = cellMap_(i - 1, colIdx).Top
              col0 = cellMap_(i - 1, colIdx).Left
              For k = row0 To i
                For l = col0 To colIdx + colSpan - 1
                  cellMap_(k, l) = cellMap_(row0, col0)
                  cellMap_(k, l).Bottom = i
                Next
              Next
            Else
              For k = colIdx To colIdx + colSpan - 1
                cellMap_(i, k).Left = colIdx
                cellMap_(i, k).Right = colIdx + colSpan - 1
                cellMap_(i, k).Top = i
                cellMap_(i, k).Bottom = i
              Next
            End If
            ' update colIdx
            colIdx = colIdx + colSpan
          Next
        End With
      Next
    End With
    
    ' remove blank rows
    i = 1
    While i <= RowCount_
      'check if all cells in the row are merged with upper cells
      flag = True
      For j = 1 To ColumnCount_
        If cellMap_(i, j).Top = i Then flag = False
      Next
      If flag Then 'blank row
        'shift array elements
        For p = i + 1 To RowCount_
            For q = 1 To ColumnCount_
              cellMap_(p - 1, q) = cellMap_(p, q)
            Next
        Next
        'update top and bottom values
        For p = 1 To RowCount_
            For q = 1 To ColumnCount_
              If cellMap_(p, q).Top >= i Then cellMap_(p, q).Top = cellMap_(p, q).Top - 1
              If cellMap_(p, q).Bottom >= i Then cellMap_(p, q).Bottom = cellMap_(p, q).Bottom - 1
            Next
        Next
        'update rowCount
        RowCount_ = RowCount_ - 1
      Else
        i = i + 1
      End If
    Wend
        
    ' set obj and IsMerge
    For i = 1 To RowCount_
      col = 1
      For j = 1 To ColumnCount_
          Set cellMap_(i, j).obj = tbl.cell(cellMap_(i, j).Top, col)
          If cellMap_(i, j).Top = i And cellMap_(i, j).Left = j Then
            cellMap_(i, j).IsMerged = False
          Else
            cellMap_(i, j).IsMerged = Top
          End If
          ' update col
          If j = cellMap_(i, j).Right Then
            col = col + 1
          End If
        Next
    Next
    
    'copy to correct size array
    ReDim cellMap2_(RowCount_, ColumnCount_)
    For i = 1 To RowCount_
      For j = 1 To ColumnCount_
          cellMap2_(i, j) = cellMap_(i, j)
        Next
    Next
    Erase cellMap_
            
    'set results
    RowCount = RowCount_
    ColumnCount = ColumnCount_
    CreateCellMap = cellMap2_
End Function




' return the element of the cell map corresponding to the cell object
Public Function GetTableCell(cellObj As Word.cell, cellMap() As TableCell) As TableCell
  For r = 1 To UBound(cellMap, 1)
    For c = 1 To UBound(cellMap, 2)
      If cellMap(r, c).obj.RowIndex = cellObj.RowIndex And cellMap(r, c).obj.ColumnIndex = cellObj.ColumnIndex Then
        GetTableCell = cellMap(r, c)
        Exit Function
      End If
    Next
  Next
  Err.Raise 9999, "", "program error in GetTableCell() In TableUtil"
End Function



' return text in the cell object (including formField, etc)
Public Function GetTableCellText(cellObj As Word.cell) As String
    GetTableCellText = ""
    'Set rng = ActiveDocument.Range(Start:=cellObj.Range.Start, End:=cellObj.Range.End - 1)
    For i = 1 To cellObj.Range.Characters.Count - 1 'remove trailing 1 elements (newline)
      Set crng = cellObj.Range.Characters(i)
      If crng.FormFields.Count > 0 Then
        For Each ff In crng.FormFields
          tmp = ""
          On Error Resume Next
          tmp = ff.Result
          On Error GoTo 0
          GetTableCellText = GetTableCellText & tmp
        Next
      Else
        GetTableCellText = GetTableCellText & crng.Text
      End If
    Next
End Function


