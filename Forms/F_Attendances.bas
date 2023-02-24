Option Compare Database


'【概要】フォーム読み込みイベント
'【作成日】2023/02/08
Private Sub Form_Load()
On Error GoTo Form_Load_Err

    'ログインチェック
    If LoginCheck(Me.Name) = False Then
       GoTo Form_Load_Exit
    End If
     
    'フォーム初期化
    Call InitForm
      
Form_Load_Err:

Form_Load_Exit:
End Sub


'【概要】フォーム初期化
'【作成日】2023/02/21
Private Sub InitForm()
On Error GoTo InitForm_Err

    'レコードソース設定
    Call SetRecordSource
     
InitForm_Err:

InitForm_Exit:
End Sub


'【概要】レコードソース設定
'【作成日】2023/02/21
Private Sub SetRecordSource()
On Error GoTo SetRecordSource_Err

    'レコードソース設定
    Me.RecordSource = GetAttendancesSQL
     
SetRecordSource_Err:

SetRecordSource_Exit:
End Sub


