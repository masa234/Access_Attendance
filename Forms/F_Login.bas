Option Compare Database

'ログインカウント
Private lngLoginCount As Long


'【概要】フォームの読み込みイベント
'【作成日】2023/02/18
Private Sub Form_Load()
On Error GoTo Form_Load_Err

    'ログインカウントを初期化する
    lngLoginCount = 0

Form_Load_Exit:
    Exit Sub
Form_Load_Err:
    GoTo Form_Load_Exit
End Sub


'【概要】ログインボタンクリックイベント
'【作成日】2023/02/02
Private Sub btnLogin_Click()
On Error GoTo btnLogin_Click_Err
    
    Dim lngLockUserID As Long
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
        GoTo btnLogin_Click_Exit
    End If
    
    '入力チェック(パスワード)
    If IsFitInText(strPassword, PASSWORD_MAX_LENGTH) = False Then
        'メッセージ
        Call MsgBox(PASSWORD_LENGTH_OVER, vbInformation, CONFIRM)
        GoTo btnLogin_Click_Exit
    End If
    
    
    'ユーザ名が存在するかどうか
    If IsExistsUserName(strUserName) = False Then
        'TODO :同じコードを2回記述している
        'メッセージボックス
        Call MsgBox(LOGIN_FAILED, vbInformation, CONFIRM)
        GoTo btnLogin_Click_Exit
    End If
        
    'ロックされているユーザ名かどうか？
    If IsLockedUserName(strUserName) = True Then
        'メッセージボックス
        Call MsgBox(USER_IS_LOCKED, vbInformation, CONFIRM)
        GoTo btnLogin_Click_Exit
    End If
    
    'ユーザが存在するかどうか
    If IsExistsUser(strUserName, strPassword) = False Then
        'ロックカウントをカウントアップ
        lngLoginCount = lngLoginCount + 1
        '指定回数以上、ログインに失敗した場合ロック
        If LOCK_COUNT <= lngLoginCount Then
            'ユーザIDを取得
            lngLockUserID = GetUserIDByUserName(strUserName)
            'ロック
            If LockUpdate(lngLockUserID, EnumIsLock.Locked) = False Then
                'メッセージボックス
                Call MsgBox(LOCK_UPDATE_FAILED, vbInformation, CONFIRM)
                GoTo btnLogin_Click_Exit
            End If
        End If
        'メッセージ
        Call MsgBox(LOGIN_FAILED, vbInformation, CONFIRM)
        GoTo btnLogin_Click_Exit
    End If
    
    'ログイン情報を保存する
    lngLoginUserID = GetUserID(strUserName, strPassword)
    
    'フォームを閉じる
    DoCmd.Close acForm, Me.Name
    
    'フォームを開く
    DoCmd.OpenForm MAIN_FORM

btnLogin_Click_Exit:
    Exit Sub
btnLogin_Click_Err:
    GoTo btnLogin_Click_Exit
End Sub

