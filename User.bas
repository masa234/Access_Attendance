Option Compare Database


'【概要】ユーザが存在するかどうか？
'【作成日】2023/02/02
Public Function IsExistsUser(ByVal strUserName As String, _
                            ByVal strPassword As String) As Boolean
On Error GoTo IsExistsUser_Err

    IsExistsUser = False
    
    Dim objCmd As ADODB.Command
    Dim objRs As ADODB.Recordset
    Dim strSQL As String
    
    'オブジェクト作成
    Call CreateCmdAndRs(objCmd, objRs)
    
    'SQL取得
    strSQL = GetUserSQL
    
    'パラム設定
    Call AddParameter(objCmd, "ユーザ名", adChar, 255, strUserName)
    Call AddParameter(objCmd, "パスワード", adChar, 255, strPassword)
    
    'SQL実行
    If ExecuteSQL(strSQL, objCmd, objRs) = False Then
        GoTo IsExistsUser_Exit
    End If
    
    'データが存在する場合、True
    If IsExistsData(objRs) = True Then
        IsExistsUser = True
    End If
    
IsExistsUser_Exit:
    'オブジェクトを破棄
    Call DisposeCmdAndRs(objCmd, objRs)
    Exit Function
IsExistsUser_Err:
    GoTo IsExistsUser_Exit
End Function


'【概要】ユーザが存在するかどうか？
'【作成日】2023/02/02
Public Function IsExistsUserName(ByVal strUserName As String) As Boolean
On Error GoTo IsExistsUserName_Err

    IsExistsUserName = False
    
    Dim strSQL As String
    Dim objCmd As ADODB.Command
    Dim objRs As ADODB.Recordset
    
    'オブジェクト作成
    Call CreateCmdAndRs(objCmd, objRs)
    
    'SQL取得
    strSQL = GetUserNameSQL
    
    'パラム設定
    Call AddParameter(objCmd, "ユーザ名", adChar, 255, strUserName)
    
    'データ取得
    If ExecuteSQL(strSQL, objCmd, objRs) = False Then
        GoTo IsExistsUserName_Exit
    End If
    
    'データが存在する場合、True
    If IsExistsData(objRs) = True Then
        IsExistsUserName = True
    End If
    
IsExistsUserName_Exit:
    'オブジェクトを破棄
    Call DisposeCmdAndRs(objCmd, objRs)
    Exit Function
IsExistsUserName_Err:
    GoTo IsExistsUserName_Exit
End Function


'【概要】ユーザID取得
'【作成日】2023/02/02
Public Function GetUserID(ByVal strUserName As String, _
                            ByVal strPassword As String) As Long
On Error GoTo GetUserID_Err
    
    Dim strSQL As String
    Dim objCmd As ADODB.Command
    Dim objRs As ADODB.Recordset
    
    'オブジェクト作成
    Call CreateCmdAndRs(objCmd, objRs)
    
    'SQL取得
    strSQL = GetUserIDSQL
    
    'パラム設定
    Call AddParameter(objCmd, "ユーザ名", adChar, 255, strUserName)
    Call AddParameter(objCmd, "パスワード", adChar, 255, strPassword)
    
    'SQL実行
    If ExecuteSQL(strSQL, objCmd, objRs) = False Then
        GoTo GetUserID_Exit
    End If
    
    With objRs
        '設定
         GetUserID = CLng(objRs(0).Value)
    End With
    
GetUserID_Exit:
    'オブジェクトを破棄
    Call DisposeCmdAndRs(objCmd, objRs)
    Exit Function
GetUserID_Err:
    GoTo GetUserID_Exit
End Function


'【概要】ユーザ名取得
'【作成日】2023/02/18
Public Function GetUserIDByUserName(ByVal strUserName As String) As Long
On Error GoTo GetUserIDByUserName_Err
    
    Dim strSQL As String
    Dim objCmd As ADODB.Command
    Dim objRs As ADODB.Recordset
    
    'オブジェクト作成
    Call CreateCmdAndRs(objCmd, objRs)
    
    'パラム設定
    Call AddParameter(objCmd, "ユーザ名", adChar, 255, strUserName)
    
    'SQL取得
    strSQL = GetUserByUserNameSQL
    
    '実行結果を取得
    If ExecuteSQL(strSQL, objCmd, objRs) = False Then
        GoTo GetUserIDByUserName_Exit
    End If
    
    With objRs
        '設定
         GetUserIDByUserName = CLng(objRs(0).Value)
    End With
    
GetUserIDByUserName_Exit:
    'オブジェクトを破棄
    Call DisposeCmdAndRs(objCmd, objRs)
    Exit Function
GetUserIDByUserName_Err:
    GoTo GetUserIDByUserName_Exit
End Function


'【概要】管理者かどうか
'【作成日】2023/02/02
Public Function IsAdmin(ByVal lngUserID As Long) As Boolean
On Error GoTo IsAdmin_Err

    IsAdmin = False
    
    Dim strSQL As String
    Dim objCmd As ADODB.Command
    Dim objRs As ADODB.Recordset
    
    'オブジェクト作成
    Call CreateCmdAndRs(objCmd, objRs)
    
    'SQL取得
    strSQL = GetAdminSQL
    
    'パラム設定
    Call AddParameter(objCmd, "ユーザID", adChar, 255, lngUserID)
    
    'SQL実行
    If ExecuteSQL(strSQL, objCmd, objRs) = False Then
        GoTo IsAdmin_Exit
    End If
    
    '管理者かどうか？
    If CInt(objRs(0).Value) = EnumUserType.Admin Then
        IsAdmin = True
    End If
    
IsAdmin_Exit:
    'オブジェクトを破棄
    Call DisposeCmdAndRs(objCmd, objRs)
    Exit Function
IsAdmin_Err:
    GoTo IsAdmin_Exit
End Function


'【概要】ロックされているかどうか
'【作成日】2023/02/18
Public Function IsLockedUserName(ByVal strUserName As String) As Boolean
On Error GoTo IsLockedUserName_Err

    IsLockedUserName = False
    
    Dim strSQL As String
    Dim objCmd As ADODB.Command
    Dim objRs As ADODB.Recordset
    
    'オブジェクト作成
    Call CreateCmdAndRs(objCmd, objRs)
    
    'SQL取得
    strSQL = GetLockedUserName
    
    'パラム設定
    Call AddParameter(objCmd, "ユーザ名", adChar, 255, strUserName)
    
    '実行結果を取得
    If ExecuteSQL(strSQL, objCmd, objRs) = False Then
        GoTo IsLockedUserName_Exit
    End If
    
    'ロックされているかどうか?
    If CInt(objRs(0).Value) = EnumIsLock.Locked Then
        IsLockedUserName = True
    End If
    
IsLockedUserName_Exit:
    'オブジェクトを破棄
    Call DisposeCmdAndRs(objCmd, objRs)
    Exit Function
IsLockedUserName_Err:
    GoTo IsLockedUserName_Exit
End Function



'【概要】ユーザ登録
'【作成日】2023/02/02
Public Function UserRegister(ByVal strUserName As String, _
                            ByVal strPassword As String) As Boolean
On Error GoTo UserRegister_Err

    UserRegister = False
    
    Dim objCmd As ADODB.Command
    Dim strSQL As String
    
    'オブジェクト作成
    Call CreateCmd(objCmd)
    
    'SQL取得
    strSQL = GetUserRegisterSQL
    
    'パラム設定
    Call AddParameter(objCmd, "ユーザ名", adChar, 255, strUserName)
    Call AddParameter(objCmd, "パスワード", adChar, 255, strPassword)
    Call AddParameter(objCmd, "管理者", adChar, 255, EnumUserType.Normal)
    Call AddParameter(objCmd, "ロックフラグ", adChar, 255, EnumIsLock.NotLock)
    
    'SQL実行
    'TODO ;動作しない（原因不明）
    If ExecuteSQL(strSQL, objCmd) = False Then
        GoTo UserRegister_Exit
    End If
    
UserRegister_Exit:
    'オブジェクトを破棄
    Call DisposeCmd(objCmd)
    Exit Function
UserRegister_Err:
    GoTo UserRegister_Exit
End Function


'【概要】ユーザ削除
'【作成日】2023/02/08
Public Function UserDelete(ByVal lngDeleteUserID As Long) As Boolean
On Error GoTo UserDelete_Err

    UserDelete = False
    
    Dim objCmd As ADODB.Command
    Dim strSQL As String
    
    'オブジェクト作成
    Call CreateCmd(objCmd)
    
    'SQL取得
    strSQL = GetUserDeleteSQL
    
    'パラム設定
    Call AddParameter(objCmd, "ユーザID", adChar, 255, lngDeleteUserID)
    
    '実行
    If ExecuteSQL(strSQL, objCmd) = False Then
        GoTo UserDelete_Exit
    End If
    
    UserDelete = True
    
UserDelete_Exit:
    'オブジェクトを破棄
    Call DisposeCmd(objCmd)
    Exit Function
UserDelete_Err:
    GoTo UserDelete_Exit
End Function


'【概要】ユーザ取得
'【作成日】2023/02/08
Public Function GetUser(ByVal lngUserID As Long) As Variant
On Error GoTo GetUser_Err
    
    Dim objCmd As ADODB.Command
    Dim objRs As ADODB.Recordset
    Dim arrRet(1) As Variant
    Dim strSQL As String
    
    'オブジェクト作成
    Call CreateCmdAndRs(objCmd, objRs)
    
    'SQL取得
    strSQL = GetUserDataSQL
    
    'パラム設定
    Call AddParameter(objCmd, "ユーザID", adChar, 255, lngUserID)
    
    'データ取得
    If ExecuteSQL(strSQL, objCmd, objRs) = False Then
        GoTo GetUser_Exit
    End If
    
    '配列に代入
    arrRet(0) = objRs(0).Value
    arrRet(1) = objRs(1).Value
    
    '配列を返却
    GetUser = arrRet
        
GetUser_Exit:
    'オブジェクトを破棄
    Call DisposeCmdAndRs(objCmd, objRs)
    Exit Function
GetUser_Err:
    GoTo GetUser_Exit
End Function


'【概要】ユーザ更新
'【作成日】2023/02/09
Public Function UserUpdate(ByVal strUserName As String, _
                            ByVal strPassword As String, _
                            ByVal lngUpdateUserID As Long) As Boolean
On Error GoTo UserUpdate_Err

    UserUpdate = False
    
    Dim objCmd As ADODB.Command
    Dim strSQL As String
    
    'オブジェクト作成
    Call CreateCmd(objCmd)
    
    'SQL取得
    strSQL = GetUserUpdateSQL
    
    'パラム設定
    Call AddParameter(objCmd, "ユーザ名", adChar, 255, strUserName)
    Call AddParameter(objCmd, "パスワード", adChar, 255, strPassword)
    Call AddParameter(objCmd, "ユーザID", adChar, 255, lngUpdateUserID)
    
    '実行
    If ExecuteSQL(strSQL, objCmd) = False Then
        GoTo UserUpdate_Exit
    End If
    
    UserUpdate = True
    
UserUpdate_Exit:
    'オブジェクトを破棄
    Call DisposeCmd(objCmd)
    Exit Function
UserUpdate_Err:
    GoTo UserUpdate_Exit
End Function


'【概要】ユーザ更新
'【作成日】2023/02/09
Public Function AuthorityUpdate(ByVal lngUpdateUserID As Long, _
                                ByVal lngUpdateUserType As Long) As Boolean
On Error GoTo AuthorityUpdate_Err

    AuthorityUpdate = False
    
    Dim objCmd As ADODB.Command
    Dim strSQL As String
    
    'オブジェクト作成
    Call CreateCmd(objCmd)
    
    'SQL取得
    strSQL = GetAdminUpdateSQL
    
    'パラム設定
    Call AddParameter(objCmd, "更新ユーザ種類", adChar, 255, lngUpdateUserType)
    Call AddParameter(objCmd, "ユーザ更新ID", adChar, 255, lngUpdateUserID)
    
    '実行
    If ExecuteSQL(strSQL, objCmd) = False Then
        GoTo AuthorityUpdate_Exit
    End If
    
    AuthorityUpdate = True
    
AuthorityUpdate_Exit:
    'オブジェクトを破棄
    Call DisposeCmd(objCmd)
    Exit Function
AuthorityUpdate_Err:
    GoTo AuthorityUpdate_Exit
End Function


'【概要】ロック状態を更新
'【作成日】2023/02/18
Public Function LockUpdate(ByVal lngUpdateUserID, _
                            ByVal lngIsLock As Long) As Boolean
On Error GoTo LockUpdate_Err

    LockUpdate = False
    
    Dim objCmd As ADODB.Command
    Dim strSQL As String
    
    'オブジェクト作成
    Call CreateCmd(objCmd)
    
    'SQL取得
    strSQL = GetLockUpdateSQL
    
    'パラム設定
    Call AddParameter(objCmd, "ロック状態", adChar, 255, lngIsLock)
    Call AddParameter(objCmd, "更新対象ID", adChar, 255, lngUpdateUserID)
    
    'SQL実行
    If ExecuteSQL(strSQL, objCmd) = False Then
        GoTo LockUpdate_Exit
    End If
    
    LockUpdate = True
    
LockUpdate_Exit:
    'オブジェクトを破棄
    Call DisposeCmd(objCmd)
    Exit Function
LockUpdate_Err:
    GoTo LockUpdate_Exit
End Function
