# VBA-Word-Table
VBA Utilities for Word Table 

This utility extract vertically/horizontally merged cells in Word tables, 
and provides the interfaces for easy access to cells in the table.

## Usage

```vb.net
  Dim T As New TableUtil
  T.Init tableObj
  For r = 1 To T.RowCount
    For c = 1 To T.ColumnCount
      Debug.Print T.Cells(r,c).Top
    Next
  Next
  For c In T.Cells
    Debug.Print c.obj.Range
  Next
```
