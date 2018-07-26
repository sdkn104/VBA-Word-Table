# VBA-Word-Table
VBA Utilities for MS Word Table 

This utility extract vertically/horizontally merged cells in Word tables, 
and provides the interfaces for easy access to cells in the table.

## Usage

```vb.net
  Dim cellMap() As TableCell
  Dim cm As TableCell
  Dim cl As Word.Cell
  Dim r As Long, c As Long

  cellMap = CreateCellMap(ActiveDocument.Tables(1), r, c)
  Debug.Print r & c

  For r = 1 To UBound(cellMap, 1)
    For c = 1 To UBound(cellMap, 2)
        cm = cellMap(r, c)
        Debug.Print "(" & cm.Top & cm.Left & ")-(" & cm.Bottom & cm.Right & ") " & cm.IsMerged & ":" & GetTableCellText(cm.obj)
    Next
  Next

  For Each cl In ActiveDocument.Tables(1).Range.Cells
    Debug.Print cl.RowIndex & cl.ColumnIndex
    cm = GetTableCell(cl, cellMap)
    Debug.Print "(" & cm.Top & cm.Left & ")-(" & cm.Bottom & cm.Right & ") " & cm.IsMerged & ":" & GetTableCellText(cm.obj)
  Next
```

### Example
<table>
  <tr>
    <td>D11</td><td colspan="2">D12</td>
  </tr>
  <tr>
   <td colspan="2" rowspan="2">D21</td><td>D23</td>
  </tr>
  <tr>
   <td>D33</td>
  </tr>
  <tr>
   <td>D41</td><td>D42</td><td>D43</td>
  </tr>
</table>

cellMap(2,1) : Top=2, Left=2, Bottom=3, Right=3, IsMerged=False

cellMap(3,2) : Top=2, Left=2, Bottom=3, Right=3, IsMerged=True
