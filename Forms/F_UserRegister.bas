Option Compare Database


'【概要】フォーム読み込みイベント
'【作成日】2023/02/09
Private Sub Form_Load()
On Error GoTo Form_Load_Err

    'ログインチェック
    If LoginCheck(Me.Name) = False Then
        GoTo Form_Load_Exit
    End If
     
Form_Load_Err:

Form_Load_Exit:
End Sub


'【概要】登録ボタンクリックイベント
'【作成日】2023/02/02
Private Sub btnUserRegister_Click()
On Error GoTo btnUserRegister_Click_Err

    Dim strUserName As String
    Dim strPassword As String

    '管理者かどうか
    If IsAdmin(lngLoginUserID) = False Then
        '管理者でない場合、終了
        '通常、ありえない挙動のためメッセージは出さない
        GoTo btnUserRegister_Click_Exit
    End If
    
    'ユーザ名
    strUserName = Nz(txtUserName.Value, vbNullString)
    'パスワード
    strPassword = Nz(txtPassword.Value, vbNullString)
    
    '入力チェック(ユーザ名)
    If IsFitInText(strUserName, USERNAME_MAX_LENGTH) = False Then
        'メッセージ
        Call MsgBox(USERNAME_LENGTH_OVER, vbInformation, CONFIRM)
        GoTo btnUserRegister_Click_Exit
    End If
    
    '入力チェック(パスワード)
    If IsFitInText(strPassword, PASSWORD_MAX_LENGTH) = False Then
        'メッセージ
        Call MsgBox(PASSWORD_LENGTH_OVER, vbInformation, CONFIRM)
        GoTo btnUserRegister_Click_Exit
    End If
    
    '登録
    If UserRegister(strUserName, strPassword) = False Then
        'メッセージ
        Call MsgBox(USER_REGISTER_FAILED, vbInformation, CONFIRM)
        GoTo btnUserRegister_Click_Exit
    End If
    
    '成功メッセージ
    Call MsgBox(USER_REGISTER_SUCCESS, vbInformation, CONFIRM)
    
btnUserRegister_Click_Err:

btnUserRegister_Click_Exit:
End Sub

