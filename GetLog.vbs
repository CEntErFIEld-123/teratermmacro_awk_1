'(1)事前設定
Option Explicit '変数を明示的に宣言

'(2)使用する変数の宣言
Dim CdPath 'カレントディレクトリのパス
Dim ParPath 'カレントディレクトリの親ディレクトのパス
Dim TTLFileName '使用するTTLファイル名
Dim VBSFileName '使用するVBSファイル名
Dim ResultFileName 'テキストファイル名
Dim UsrInputStr '使用者が入力した値
Dim ReDate '使用者が入力した値をリネームした値
Dim StartDate 'TTLファイルへ引き渡す日時の値
Dim EndDate 'TTLファイルへ引き渡す日時の値
Dim Fso 'ファイルシステムオブジェクト用に使用
Dim CheckStrError '入力された値のチェックで使用
Dim CheckStrNum '数字の判定に使用
Dim CheckMsg 'メッセージの戻り確認に使用
Dim i, j 'カウンター

'(3)使用するttlとvbsファイルが存在しているか確認する。
Set Fso = CreateObject("Scripting.FileSystemObject")
CdPath = Fso.GetAbsolutePathName(".") 'カレントディレクトリのパスを取得
ParPath = Fso.GetAbsolutePathName("..") 'カレントディレクトリの親ディレクトのパスを取得
TTLFileName = "test.ttl"
VBSFileName = "test.vbs"
ResultFileName = "result.txt"
 'ファイルの存在チェック
If Fso.FileExists(CdPath & "\" & TTLFileName) = False Then
  MsgBox "フォルダ「" & CdPath & "」に使用するファイル(" & TTLFileName & ")が存在しません。該当フォルダを確認してください。", vbCritical,"使用するファイルがありません"
  Set FSO = Nothing
  WScript.Quit
ElseIf Fso.FileExists(CdPath & "\" & VBSFileName) = False Then 
  MsgBox "フォルダ「" & CdPath & "」に使用するファイル(" & VBSFileName & ")が存在しません。該当フォルダを確認してください。", vbCritical,"使用するファイルがありません"
  Set FSO = Nothing
  WScript.Quit
ElseIf Fso.FileExists(ParPath & "\結果\" & ResultFileName) = True Then
  MsgBox "フォルダ「" & ParPath & "\結果」にテラタームマクロがログを書き込むファイル(" & ResultFileName & ")が存在します。" & vbCr & _
  "退避漏れの可能性があります。ファイルを削除するか退避をしてください。", vbCritical,"マクロがログを書き込むファイルが残っています"
  Set FSO = Nothing
  WScript.Quit
End If
Set FSO = Nothing

'(4)日付の入力処理
CheckStrError = "" '念のため
UsrInputStr = ""
Do Until CheckStrError = "No Error"
  '使用者へ入力要求
 UsrInputStr = CStr(InputBox("serverinfoの" & vbCr & _ 
 "「path」より入力された日時の10分前～5分後までの間に書き込まれた" & vbCr & "メッセージを取得します。" & vbCr & _
  "なお、入力する日時は<< YYYY/MM/DD hh:mm:ss >>の形式でお願いします。" & vbCr & _
  "※なにも入力せずに「OK」を押すと終了できます。" , "データを取得する日時を入力してください。", UsrInputStr))
  CheckStrError = ""
  If UsrInputStr = "" Then
    msgbox "ツールを終了します。",vbInformation
    WScript.Quit
  End if
  '(5)入力された値が正しいか確認する。
  '①文字数の確認
  If CheckStrError = "" Then CheckStrError = CheckLen(UsrInputStr)
  '②「/」の確認
  If CheckStrError = "" Then CheckStrError = CheckSla(UsrInputStr)
  '③スペースの確認
  If CheckStrError = "" Then CheckStrError = CheckSpa(UsrInputStr)
  '④「:」の確認
  If CheckStrError = "" Then CheckStrError = CheckCol(UsrInputStr)
  '⑤数値の判定
  If CheckStrError = "" Then CheckStrError = CheckNum(UsrInputStr)
  '⑥日付の整合性の確認
  If CheckStrError = "" Then CheckStrError = CheckDate(UsrInputStr)
   '年月が本日と違っていた場合の確認
  If CheckStrError = "No Error" Then
    If FormatDateTime(Now, 1) <> FormatDateTime(UsrInputStr, 1) Then
      CheckMsg = MsgBox("入力された年月日(" & FormatDateTime(UsrInputStr, 1) & ")と" & vbCr & _
      "本日の年月日(" & FormatDateTime(Now, 1) & ")に乖離があります。。" & vbCr & "本日の年月日が異なっている状態で実行しますか？", vbyesno + vbExclamation,"年月日に乖離があります")
      If CheckMsg = vbNo then
        CheckStrError = ""
        msgbox "再度入力をお願いします。",vbInformation
      End If
    End if
  End if
Loop

'(6)入力された値から10分前を計算してログ収集開始日時のリネーム
ReDate = DateAdd("n", -10, UsrInputStr)
ReDate = Replace(ReDate, "/", "-")
StartDate = Replace(ReDate, " ", "T")
msgbox StartDate

'(7)入力された値から5分後の計算をしてログ収集終了日時をリネーム。
ReDate = DateAdd("n", 5, UsrInputStr)
ReDate = Replace(ReDate, "/", "-")
EndDate = Replace(ReDate, " ", "T")
msgbox EndDate

'(8)テラタームマクロへ()と()の値を引き渡す。


'VBSでの処理は終了
'==================================================
'========<以下、サブルーチンです>==================
'(5)入力された値が正しいか確認する。
'①文字数の確認
Function CheckLen(UsrInputStr)
  If Len(UsrInputStr) > 19 Then
    CheckLen = "Error"
    MsgBox "入力された日付の文字数が多いです。もう一度入力してください。", vbCritical,"入力エラー"
  ElseIf  Len(UsrInputStr) < 19 Then
    CheckLen = "Error"
    MsgBox "入力された日付の文字数が少ないです。もう一度入力してください。", vbCritical,"入力エラー"
  End if
End Function

'②「/」の確認
Function CheckSla(UsrInputStr)
  j = 0
  For i = 1 To 2
    j = InStr(j + 1, UsrInputStr, "/")
  Next
  If j <> 8 Then
    CheckSla = "Error"
    MsgBox "入力された日付の年月日の「/」が正しく入力されていません。半角でもう一度入力してください。", vbCritical,"入力エラー"
  End if
End Function

'③スペースの確認
Function CheckSpa(UsrInputStr)
  j = 0
  j = InStr(UsrInputStr, " ")
  If j <> 11 Then
   CheckSpa = "Error"
   MsgBox "入力された日付の年月日と時間の間にスペースが正しく入力されていません。", vbCritical,"入力エラー"
  End if
End Function

'④「:」の確認
Function CheckCol(UsrInputStr)
  j = 0
  For i = 1 To 2
    j = InStr(j + 1, UsrInputStr, ":")
  Next
  If j <> 17 Then
    CheckCol = "Error"
    MsgBox "入力された日付の時間の「:」が正しく入力されていません。半角でもう一度入力してください。", vbCritical,"入力エラー"
  End if
End Function

'⑤数値の判定
Function CheckNum(UsrInputStr)
  CheckStrNum = Replace(UsrInputStr, "/", "")
  CheckStrNum = Replace(CheckStrNum, " ", "")
  CheckStrNum = Replace(CheckStrNum, ":", "")
  If IsNumeric(CheckStrNum) = False Or Len(CheckStrNum) <> 14 Then '数字以外が含まれているか数字が足りない場合エラー
      CheckNum = "Error"
      MsgBox "入力された日付に数字以外が含まれています。もう一度入力してください。", vbCritical,"入力エラー"
  End if
End Function

'⑥日付の整合性の確認
Function CheckDate(UsrInputStr)
  If IsDate(UsrInputStr) = true Then
    CheckDate = "No Error"
  Else
    CheckDate = "Error"
    MsgBox "入力された日付に誤りがあります。もう一度入力してください。", vbCritical,"入力エラー"
  End if
End Function
