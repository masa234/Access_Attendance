Option Compare Database



'【概要】フォーム読み込みイベント
'【作成日】2023/02/02
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


'【概要】登録ボタンクリックイベント
'【作成日】2023/02/02
Private Function InitForm() As Boolean
On Error GoTo InitForm_Err

    InitForm = False
    
    '管理者でない場合
     If IsAdmin(lngLoginUserID) = False Then
        'ボタンを非表示にする
        'ユーザ登録
        Me.btnUserRegister.Visible = False
        'ユーザ一覧
        Me.btnUsers.Visible = False
        'ロックされているユーザ一覧
        Me.btnLockUsers.Visible = False
     End If
     
     InitForm = True
    
InitForm_Err:

InitForm_Exit:
End Function


'【概要】登録ボタンクリックイベント
'【作成日】2023/02/02
Private Sub btnUserRegister_Click()
On Error GoTo btnUserRegister_Click_Err

    '自画面を閉じる
    DoCmd.Close acForm, Me.Name
    
    '登録画面を開く
    DoCmd.OpenForm USER_REGISTER_FORM
    
btnUserRegister_Click_Err:

btnUserRegister_Click_Exit:
End Sub


'【概要】ユーザ一覧ボタンクリックイベント
'【作成日】2023/02/02
Private Sub btnUsers_Click()

    '自画面を閉じる
    DoCmd.Close acForm, Me.Name
    
    'ユーザ一覧画面を開く
    DoCmd.OpenForm USERS_FORM

End Sub


'【概要】ユーザ更新ボタンクリックイベント
'【作成日】2023/02/09
Private Sub btnUserUpdate_Click()

    '自画面を閉じる
    DoCmd.Close acForm, Me.Name
    
    'ユーザ更新画面を開く
    DoCmd.OpenForm USER_UPDATE_FORM

End Sub


'【概要】ロックユーザ一覧ボタンクリックイベント
'【作成日】2023/02/18
Private Sub btnLockUsers_Click()
On Error GoTo btnLockUsers_Click_Err

    '自画面を閉じる
    DoCmd.Close acForm, Me.Name
    
    'ユーザ更新画面を開く
    DoCmd.OpenForm LOCK_USERS_FORM
    
btnLockUsers_Click_Err:

btnLockUsers_Click_Exit:
End Sub


'【概要】出勤打刻ボタンクリックイベント
'【作成日】2023/02/18
Private Sub btnAttendanceEngraving_Click()
On Error GoTo btnAttendanceEngraving_Click_Err

    '出勤打刻
    If AttendanceInsert(loginuserid, EnumAttendanceType.Attendance) = False Then
        'メッセージボックス
        Call MsgBox(ATTENDANCE_INSERT_FAILED, vbInformation, CONFIRM)
        GoTo btnAttendanceEngraving_Click_Exit
    End If
    
    '成功メッセージ
    Call MsgBox(ATTENDANCE_INSERT_SUCCESS, vbInformation, CONFIRM)
    
btnAttendanceEngraving_Click_Err:

btnAttendanceEngraving_Click_Exit:
End Sub


'【概要】退勤打刻ボタンクリックイベント
'【作成日】2023/02/21
Private Sub btnLeavingEngraving_Click()
On Error GoTo btnLeavingEngraving_Click_Err

    '出勤打刻
    If AttendanceInsert(loginuserid, EnumAttendanceType.Leaving) = False Then
        'メッセージボックス
        Call MsgBox(ATTENDANCE_INSERT_FAILED, vbInformation, CONFIRM)
        GoTo btnLeavingEngraving_Click_Exit
    End If
    
    '成功メッセージ
    Call MsgBox(ATTENDANCE_INSERT_SUCCESS, vbInformation, CONFIRM)
    
btnLeavingEngraving_Click_Err:

btnLeavingEngraving_Click_Exit:
End Sub


'【概要】勤怠一覧ボタンクリックイベント
'【作成日】2023/02/21
Private Sub btnAttendances_Click()
On Error GoTo btnAttendances_Click_Err

    '自画面を閉じる
    DoCmd.Close acForm, Me.Name
    
    '勤怠一覧画面を表示
    DoCmd.OpenForm ATTENDANCES_FORM
    
btnAttendances_Click_Err:

btnAttendances_Click_Exit:
End Sub


'【概要】勤怠出力ボタンクリックイベント
'【作成日】2023/02/21
Private Sub btnAttendanceOutput_Click()
On Error GoTo btnAttendanceOutput_Click_Err

    Dim strCSVOutputPath As String

    'CSVファイルパス
    strCSVOutputPath = CurrentProject.Path
    
    'CSV出力
    If CSVOutput(strCSVOutputPath, CSV_FILE_NAME, ATTENDANCE_TABLE_NAME) = False Then
        'メッセージボックス
        Call MsgBox(CSV_OUTPUT_FAILED, vbInformation, CONFIRM)
    End If
        
    '成功メッセージ
    Call MsgBox(CSV_OUTPUT_SUCCESS, vbInformation, CONFIRM)
    
btnAttendanceOutput_Click_Err:

btnAttendanceOutput_Click_Exit:
End Sub
