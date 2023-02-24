Option Compare Database


'【概要】フォーム読み込みイベント
'【作成日】2023/02/09
Private Sub Form_Load()
On Error GoTo Form_Load_Err

    'ログインチェック
     If LoginCheck(Me.Name) = False Then
        GoTo Form_Load_Exit
     End If

    'フォーム初期化
    If InitForm() = False Then
        GoTo Form_Load_Exit
    End If
    
Form_Load_Err:

Form_Load_Exit:
End Sub


'【概要】フォーム初期化
'【作成日】2023/02/09
Private Function InitForm() As Boolean
On Error GoTo InitForm_Err

    InitForm = False

    'ユーザ情報を取得
    arrUser = GetUser(lngLoginUserID)
    
    '画面に設定
    'ユーザ名
    Me.txtUserName = arrUser(0)
    'パスワード
    Me.txtPassword = arrUser(1)
     
    InitForm = True

InitForm_Err:

InitForm_Exit:
End Function


'【概要】更新ボタンクリックイベント
'【作成日】2023/02/08
Private Sub btnUserUpdate_Click()
On Error GoTo btnUserUpdate_Click_Err

    Dim strUserName As String
    Dim strPassword As String

    'ユーザ名
    strUserName = Nz(txtUserName.Value, vbNullString)
    'パスワード
    strPassword = Nz(txtPassword.Value, vbNullString)

    '入力チェック(ユーザ名)
    If IsFitInText(strUserName, USERNAME_MAX_LENGTH) = False Then
        'メッセージ
        Call MsgBox(USERNAME_LENGTH_OVER, vbInformation, CONFIRM)
        GoTo btnUserUpdate_Click_Exit
    End If
    
    '入力チェック(パスワード)
    If IsFitInText(strPassword, PASSWORD_MAX_LENGTH) = False Then
        'メッセージ
        Call MsgBox(PASSWORD_LENGTH_OVER, vbInformation, CONFIRM)
        GoTo btnUserUpdate_Click_Exit
    End If
    
    '更新
    If UserUpdate(strUserName, strPassword, lngLoginUserID) = False Then
        'メッセージ
        Call MsgBox(USER_UPDATE_FAILED, vbInformation, CONFIRM)
    End If
    
    '成功メッセージ
    Call MsgBox(USER_UPDATE_SUCCESS, vbInformation, CONFIRM)

btnUserUpdate_Click_Err:

btnUserUpdate_Click_Exit:
End Sub

