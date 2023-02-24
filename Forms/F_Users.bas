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
     If InitForm = False Then
        GoTo Form_Load_Exit
     End If
     
Form_Load_Err:

Form_Load_Exit:
End Sub


'【概要】削除ボタンクリックイベント
'【作成日】2023/02/08
Private Sub btnUserDelete_Click()
On Error GoTo btnUserDelete_Click_Err

    Dim lngDeleteUserID As Long

    '削除対象ID
    lngDeleteUserID = CLng(Me.txtID.Value)
    
    '削除を行う
    If UserDelete(lngDeleteUserID) = False Then
        'メッセージボックス
        Call MsgBox(USER_DELETE_FAILED, vbInformation, CONFIRM)
        GoTo btnUserDelete_Click_Exit
    End If
    
    '画面を更新
    Me.Requery
    
btnUserDelete_Click_Err:

btnUserDelete_Click_Exit:
End Sub


'【概要】管理者にするボタンクリックイベント
'【作成日】2023/02/09
Private Sub btnUpdateAdmin_Click()
On Error GoTo btnUpdateAdmin_Click_Err

    Dim lngUpdateUserID As Long
    
    'ログインユーザが管理者かどうか
    If IsAdmin(lngLoginUserID) = False Then
        '管理者でない場合、終了（メッセージは通知しない）
        GoTo btnUpdateAdmin_Click_Exit
    End If

    '更新対象ユーザID
    lngUpdateUserID = CLng(Me.txtID.Value)

    '管理者かどうか
    If IsAdmin(lngUpdateUserID) = True Then
        '既に管理者の場合、終了
        Call MsgBox(USER_IS_ADMIN, vbInformation, CONFIRM)
        GoTo btnUpdateAdmin_Click_Exit
    End If
        
    '管理者に更新
    If AuthorityUpdate(lngUpdateUserID, EnumUserType.Admin) = False Then
        'メッセージボックス
        Call MsgBox(AUTHORITY_UPDATE_FAILED, vbInformation, CONFIRM)
        GoTo btnUpdateAdmin_Click_Exit
    End If
    
    '成功メッセージ
    Call MsgBox(AUTHORITY_UPDATE_SUCCESS, vbInformation, CONFIRM)
    
btnUpdateAdmin_Click_Err:

btnUpdateAdmin_Click_Exit:
End Sub


'【概要】一般ユーザにするボタンクリックイベント
'【作成日】2023/02/17
Private Sub btnUpdateNormal_Click()
On Error GoTo btnUpdateNormal_Click_Err
    
    Dim lngUpdateUserID As Long
    
    'ログインユーザが管理者かどうか
    If IsAdmin(lngLoginUserID) = False Then
        GoTo btnUpdateNormal_Click_Exit
    End If
    
    '更新対象ID
    lngUpdateUserID = CLng(Me.txtID)
    
    '既に一般ユーザかどうか
    If IsAdmin(lngUpdateUserID) = False Then
        '既に一般ユーザの場合、終了
        Call MsgBox(USER_IS_NORMAL, vbInformation, CONFIRM)
        GoTo btnUpdateNormal_Click_Exit
        
    End If
        
    '一般ユーザに更新
    If AuthorityUpdate(lngUpdateUserID, EnumUserType.Normal) = False Then
        'メッセージボックス
        Call MsgBox(AUTHORITY_UPDATE_FAILED, vbInformation, CONFIRM)
        GoTo btnUpdateNormal_Click_Exit
    End If
        
    '成功メッセージ
    Call MsgBox(AUTHORITY_UPDATE_SUCCESS, vbInformation, CONFIRM)
    
btnUpdateNormal_Click_Err:

btnUpdateNormal_Click_Exit:
End Sub


'【概要】レコードソース設定
'【作成日】2023/02/09
Private Sub SetRecordSource()
On Error GoTo SetRecordSource_Err

    '設定
    Me.RecordSource = GetUsersSQL
    
SetRecordSource_Err:

SetRecordSource_Exit:
End Sub


'【概要】ボタンの表示非表示を設定
'【作成日】2023/02/17
Private Sub SetButtonVisible()
On Error GoTo SetButtonVisible_Err

    '管理者でない場合
    If IsAdmin(lngLoginUserID) = False Then
        'ボタンを非表示にする
        '管理者にするボタン
        Me.btnUpdateAdmin.Visible = False
        '一般ユーザにするボタン
        Me.btnUpdateNormal.Visible = False
     End If
    
SetButtonVisible_Err:

SetButtonVisible_Exit:
End Sub


'【概要】フォーム初期化
'【作成日】2023/02/09
Private Function InitForm() As Boolean
On Error GoTo InitForm_Err

    InitForm = False
    
    'ボタンの表示非表示を設定
    Call SetButtonVisible
     
    'レコードソース設定
    Call SetRecordSource
     
    InitForm = True

InitForm_Err:

InitForm_Exit:
End Function


