'======================================================
'⑤レプリケーションのIPアドレスでフィルター
 '配列の作成
arr = Array("A", "B", "C")
 '作成した配列によりフィルターをかけて格納する配列の大きさを求める
ws.Range("A1").AutoFilter field:=3, Criteria1:=arr, Operator:=xlFilterValues
'⑥数を取得し、必要な数を減算及び'⑦合算
i = WorksheetFunction.Subtotal(3, Range("A1").CurrentRegion.Columns(1)) - 2
'配列の大きさを再設定
ReDim setarr(i)
ws.Range("A1").AutoFilter
'対象であれば配列へ格納
Do Until ws.Cells(j, 1) = ""
  '⑧-1 '<<<<<<既存
  '⑧-2
  For i = 0 To UBound(arr)
    '⑧-2-1
    If ws.Cells(j, 2) = arr(i) Then
      '⑧-2-1-1
      setarr(k) = ws.Cells(j, 1)
      ws.Cells(j, 4) = arr(i) & "なので該当です"
      k = k + 1 '配列の添え字ようカウンター変数の加算
      Exit For
    End If
  Next i
  j = j + 1
Loop
'======================================================
