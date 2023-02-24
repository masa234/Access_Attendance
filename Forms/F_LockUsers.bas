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


'【概要】ロック解除ボタンクリックイベント
'【作成日】2023/02/18
Private Sub btnLockRelease_Click()
On Error GoTo btnLockRelease_Click_Err

    '管理者でない場合、終了
    If IsAdmin(lngLoginUserID) = False Then
        GoTo btnLockRelease_Click_Exit
    End If
    
    'ロック解除
    If LockUpdate(Me.txtID.Value, EnumIsLock.NotLock) = False Then
        'メッセージボックス
        Call MsgBox(LOCK_UPDATE_FAILED, vbInformation, CONFIRM)
        GoTo btnLockRelease_Click_Exit
    End If
    
    'レコードソース設定
    Call SetRecordSource

    '成功メッセージ
    Call MsgBox(LOCK_UPDATE_SUCCESS, vbInformation, CONFIRM)
     
btnLockRelease_Click_Err:

btnLockRelease_Click_Exit:
End Sub


'【概要】フォーム初期化
'【作成日】2023/02/18
Private Sub InitForm()
On Error GoTo InitForm_Err

    'レコードソース設定
    Call SetRecordSource
     
InitForm_Err:

InitForm_Exit:
End Sub


'【概要】レコードソース設定
'【作成日】2023/02/18
Private Sub SetRecordSource()
On Error GoTo SetRecordSource_Err

    'レコードソース設定
    'TOOD : パラメータの渡し方わからない
    Me.RecordSource = GetLockedUsersSQL
     
SetRecordSource_Err:

SetRecordSource_Exit:
End Sub


