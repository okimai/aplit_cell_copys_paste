'--------------------------------------------------------------------
' 連続していないセルの情報を記録する.
' ※記録された情報は後ほどのペーストに役に立つ.
' また、見た目上では、指定のセルがコピーされるように見える.
' @author yo
'--------------------------------------------------------------------
Sub Copy()
  
  On Error GoTo ErrorHandler
  
  Range("1:4").Clear                       '1～4行目をクリア
  
  Range("A1").Value = Selection.Count      'セル数
  
  Dim index As Long                        'ループ回数
  index = 1
  
  '連続していないセルの情報を記録する
  For i = 1 To Selection.Areas.Count                                   '各範囲をループする
    For j = 1 To Selection.Areas(i).Count                              '各セルをループする
    
      Set cell0 = Selection.Areas(i)(j)
      
      '可視セルのみコピーする
      If cell0.EntireColumn.Hidden = False Then
      
        cells(2, index).Value = cell0.row - Selection(1).row           '2行目: 一つ目のセルよりの行offset
        cells(3, index).Value = cell0.column - Selection(1).column     '3行目: 一つ目のセルよりの列offset
        cells(4, index).Value = cell0.Value                            '4行目: セルの値
        index = index + 1
        
      End If
      
    Next j
  Next i
  
  'セルのコピー状態を付ける
  Selection.SpecialCells(xlCellTypeVisible).Copy
  
  Exit Sub

ErrorHandler:
    '-- 例外処理
     MsgBox Err.Description

End Sub

'--------------------------------------------------------------------
' 記録された情報に基づいて、ペーストする.
' また、見た目上では、指定のセルがコピーされるように見える.
'--------------------------------------------------------------------
Sub Paste()
  
  'セルのコピー状態を解除する
  Application.CutCopyMode = False
  
  Dim firstCell As Range    '選択されたセル一つ目のセル
  Dim num As Long           'コピーされたセルの数量
  Dim rowOffset As Long     '一つ目のセルよりの行offset
  Dim columnOffset As Long  '一つ目のセルよりの列offset
  Dim val As String         'セルの値
  
  Set firstCell = Selection(1)
  num = Range("A1").Value
  
  '記録された情報に基づいて、ペーストする.
  For i = 1 To num
    rowOffset = cells(2, i).Value
    columnOffset = cells(3, i).Value
    val = cells(4, i).Value
    firstCell.Offset(rowOffset, columnOffset).Value = val
  Next

End Sub
