'
' Utilities for Word Table
'

'
'  Dim T As New TableUtil
'  T.Init tableObj
'  For r = 1 To T.RowCount
'    For c = 1 To T.ColumnCount
'      Debug.Print T.Cells(r,c).Top
'    Next
'  Next
'  For c In T.Cells
'    Debug.Print c.obj.Range
'  Next
'

' Type Cell
Public Type TableCell
    Left As Long  'Leftmost column index
    Right As Long  'Rightmost column index
    Top As Long  'Top row index
    Bottom As Long 'Bottom row index
    obj As Word.Cell 'Table Object
    IsMerged As Boolean 'True if the cell is merged with another cell and not the top-left cell in the merged cells
End Type

'Private Members
Private Cell_() As TableCell
Private Cells_ As Collection
Private RowCount_ As Long
Private ColumnCount_ As Long


'Constructor
Private Sub Class_Initialize()
End Sub

'Destructor
Private Sub Class_Terminate()
  Erase Cell_
  Set Cells_ = Nothing
End Sub


'Initialization with Word Table class
Public Sub Init(tbl As Table)
  Dim d As Object
  Dim colspan As Long, colIdx As Long
  Dim i As Long, j As Long, k As Long, l As Long
     
  RowCount_ = tbl.Rows.Count
  ColumnCount_ = tbl.Rows(1).Cells.Count
  ReDim Cell_(RowCount_, ColumnCount_)
  
  colIdx = 1
  Set d = CreateObject("MSXML2.DOMDocument")
  If d.LoadXML(tbl.Range.XML) Then
    With d.SelectNodes("/w:wordDocument/w:body/wx:sect/w:tbl/w:tr")
      For i = 1 To .Length
        With .Item(i - 1).SelectNodes("w:tc")
          For j = 1 To .Length
            'get colspan
            colspan = 1
            If .Item(j - 1).SelectNodes("w:tcPr/w:gridSpan").Length > 0 Then ' horizontally merged
              colspan = CLng(.Item(j - 1).SelectNodes("w:tcPr/w:gridSpan").Item(0).Attributes(0).Text)
            End If
            'update ColumnCount_
            If colIdx > ColmunCount_ Then
              ColumnCount_ = colIdx
              ReDim Preserve Cell_(RowCount_, ColumnCount_)
            End If
            'set cell
            Set obj = Nothing
            On Error Resume Next
            Set obj = tbl.Cell(i, j)
            On Error GoTo 0
            If IsNothing(obj) Then ' vertically merged
              colspan = Cell_(i - 1, colIdx).Left - Cell_(i - 1, colIdx).Right + 1
              Cell_(i, colIdx) = Cell_(i - 1, colIdx)
              Cell_(i, colIdx).IsMerged = True
              For k = Cell_(i, colIdx).Top To i
                For l = colIdx To colIdx + colspan - 1
                  Cell_(k, l).Bottom = i
                Next
              Next
            Else
              For k = colIdx To colIdx + colspan - 1
                Cell_(i, k).Left = colIdx
                Cell_(i, k).Right = colIdx + colspan - 1
                Cell_(i, k).Top = i
                Cell_(i, k).Bottom = i
                Cell_(i, k).obj = obj
                Cell_(i, k).IsMerged = IIf(k = colIdx, False, True)
              Next
            End If
            ' update colIdx
            colIdx = colIdx + colspan
          Next
        End With
      Next
    End With
  End If
  
  ' Set Cells_ collection
  Set Cells_ = New Collection
  For i = 1 To RowCount_
    For j = 1 To ColumnCount_
      If Not Cell_(i, j).IsMerged Then
        Cells_.Add Cells_(i, j).obj
      End If
    Next
  Next
End Sub

' Return TableCell of cell(row, col)
Property Get Cells(row As Long, col As Long) As TableCell
  Set Cells = Cell_(row, col)
End Property

' Return Collection of all cells
Property Get Cells() As Collection
  Set Cells = Cells_
End Property

' Return Number of rows
Property Get RowCount() As Long
  RowCount = RowCount_
End Property

' Return Number of columns
Property Get ColumnCount() As Long
  ColumnCount = ColumnCount_
End Property

