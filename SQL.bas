Option Compare Database


'【概要】 ユーザ取得用SQLを取得
'【作成日】2023/02/02
Public Function GetUserSQL() As String
On Error GoTo GetUserSQL_Err
    
    Dim strSQL As String
    
    'ユーザーテーブルにデータを挿入するSQL
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " Id,"
    strSQL = strSQL & " UserName,"
    strSQL = strSQL & " Password, "
    strSQL = strSQL & " Admin "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " Users "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " UserName = ?"
    strSQL = strSQL & "AND "
    strSQL = strSQL & " Password = ?"
    
    
    GetUserSQL = strSQL
    
GetUserSQL_Exit:
    Exit Function
GetUserSQL_Err:
    GoTo GetUserSQL_Exit
End Function


'【概要】 ユーザ名取得用SQLを取得
'【作成日】2023/02/18
Public Function GetUserNameSQL() As String
On Error GoTo GetUserNameSQL_Err
    
    Dim strSQL As String
    
    'ユーザーテーブルにデータを挿入するSQL
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " UserName "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " Users "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " UserName = ?"
    
    GetUserNameSQL = strSQL
    
GetUserNameSQL_Exit:
    Exit Function
GetUserNameSQL_Err:
    GoTo GetUserNameSQL_Exit
End Function



'【概要】 ユーザ取得用SQLを取得
'【作成日】2023/02/18
'TODO :関数名改善の余地あり
Public Function GetUserByUserNameSQL() As String
On Error GoTo GetUserByUserNameSQL_Err
    
    Dim strSQL As String
    
    'ユーザーテーブルにデータを挿入するSQL
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " Id, "
    strSQL = strSQL & " UserName, "
    strSQL = strSQL & " Password "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " Users "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " UserName = ?"
    
    GetUserByUserNameSQL = strSQL
    
GetUserByUserNameSQL_Exit:
    Exit Function
GetUserByUserNameSQL_Err:
    GoTo GetUserByUserNameSQL_Exit
End Function


'【概要】 ユーザID取得用SQLを取得
'【作成日】2023/02/02
Public Function GetUserIDSQL() As String
On Error GoTo GetUserIDSQL_Err
    
    Dim strSQL As String
    
    'ユーザーテーブルにデータを挿入するSQL
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " Id "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " Users "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & " UserName = ? "
    strSQL = strSQL & "AND "
    strSQL = strSQL & " Password = ? "
    
    GetUserIDSQL = strSQL
    
GetUserIDSQL_Exit:
    Exit Function
GetUserIDSQL_Err:
    GoTo GetUserIDSQL_Exit
End Function


'【概要】 ユーザ登録用SQLを取得
'【作成日】2023/02/02
Public Function GetUserRegisterSQL() As String
On Error GoTo GetUserRegisterSQL_Err
    
    Dim strSQL As String
    
    'ユーザーテーブルにデータを挿入するSQL
    strSQL = vbNullString
    strSQL = strSQL & "INSERT INTO Users"
    strSQL = strSQL & "("
    strSQL = strSQL & " UserName,"
    strSQL = strSQL & " Password, "
    strSQL = strSQL & " Admin, "
    strSQL = strSQL & " LockFlag "
    strSQL = strSQL & ")"
    strSQL = strSQL & "VALUES "
    strSQL = strSQL & " (?, "
    strSQL = strSQL & "  ?, "
    strSQL = strSQL & "  ?, "
    strSQL = strSQL & "  ?) "
    
    GetUserRegisterSQL = strSQL
    
GetUserRegisterSQL_Exit:
    Exit Function
GetUserRegisterSQL_Err:
    GoTo GetUserRegisterSQL_Exit
End Function


'【概要】管理者情報取得用SQLを取得
'【作成日】2023/02/02
Public Function GetAdminSQL() As String
On Error GoTo GetAdminSQL_Err
    
    Dim strSQL As String
    
    'ユーザーテーブルにデータを挿入するSQL
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " Admin "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " Users "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & " Id = ? "
    
    GetAdminSQL = strSQL
    
GetAdminSQL_Exit:
    Exit Function
GetAdminSQL_Err:
    GoTo GetAdminSQL_Exit
End Function


'【概要】ロック状態取得用SQLを取得
'【作成日】2023/02/18
Public Function GetLockedUserName() As String
On Error GoTo GetLockedUserName_Err
    
    Dim strSQL As String
    
    'ユーザーテーブルにデータを挿入するSQL
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " LockFlag "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " Users "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & " UserName = ? "
    
    GetLockedUserName = strSQL
    
GetLockedUserName_Exit:
    Exit Function
GetLockedUserName_Err:
    GoTo GetLockedUserName_Exit
End Function


'【概要】ユーザ情報取得用SQLを取得
'【作成日】2023/02/09
Public Function GetUserDataSQL(Optional strPassword As String) As String
On Error GoTo GetUserDataSQL_Err
    
    Dim strSQL As String
    
    'ユーザーテーブルにデータを挿入するSQL
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " UserName, "
    strSQL = strSQL & " Password "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " Users "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & " Id = ? "
    
    GetUserDataSQL = strSQL
    
GetUserDataSQL_Exit:
    Exit Function
GetUserDataSQL_Err:
    GoTo GetUserDataSQL_Exit
End Function


'【概要】ユーザ情報取得用SQLを取得
'【作成日】2023/02/08
Public Function GetUsersSQL() As String
On Error GoTo GetUsersSQL_Err
    
    Dim strSQL As String
    
    'ユーザーテーブルにデータを挿入するSQL
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " Id, "
    strSQL = strSQL & " UserName "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " Users "
    strSQL = strSQL & "ORDER BY "
    strSQL = strSQL & " Id ASC "
    
    GetUsersSQL = strSQL
    
GetUsersSQL_Exit:
    Exit Function
GetUsersSQL_Err:
    GoTo GetUsersSQL_Exit
End Function


'【概要】ロックされているユーザ情報取得用SQLを取得
'【作成日】2023/02/18
Public Function GetLockedUsersSQL() As String
On Error GoTo GetLockedUsersSQL_Err
    
    Dim strSQL As String
    
    'ユーザーテーブルにデータを挿入するSQL
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " Id, "
    strSQL = strSQL & " UserName "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " Users "
    strSQL = strSQL & "WHERE "
    'TODO: ほかに方法があるかもしれない
    strSQL = strSQL & " LockFlag = 1 "
    strSQL = strSQL & "ORDER BY "
    strSQL = strSQL & " Id ASC "
    
    GetLockedUsersSQL = strSQL
    
GetLockedUsersSQL_Exit:
    Exit Function
GetLockedUsersSQL_Err:
    GoTo GetLockedUsersSQL_Exit
End Function


'【概要】ユーザ削除用SQLを取得
'【作成日】2023/02/09
Public Function GetUserDeleteSQL() As String
On Error GoTo GetUserDeleteSQL_Err
    
    Dim strSQL As String
    
    'ユーザーテーブルにデータを挿入するSQL
    strSQL = vbNullString
    strSQL = strSQL & "DELETE "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " Users "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " Id = ?"
    
    GetUserDeleteSQL = strSQL
    
GetUserDeleteSQL_Exit:
    Exit Function
GetUserDeleteSQL_Err:
    GoTo GetUserDeleteSQL_Exit
End Function


'【概要】ユーザ更新用SQLを取得
'【作成日】2023/02/09
Public Function GetUserUpdateSQL() As String
On Error GoTo GetUserUpdateSQL_Err
    
    Dim strSQL As String
    
    'ユーザーテーブルにデータを挿入するSQL
    strSQL = vbNullString
    strSQL = strSQL & "UPDATE "
    strSQL = strSQL & " Users "
    strSQL = strSQL & "SET"
    strSQL = strSQL & " UserName = ?"
    strSQL = strSQL & "AND "
    strSQL = strSQL & " Password = ?"
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & " Id = ?"
    
    GetUserUpdateSQL = strSQL
    
GetUserUpdateSQL_Exit:
    Exit Function
GetUserUpdateSQL_Err:
    GoTo GetUserUpdateSQL_Exit
End Function


'【概要】管理者権限更新用SQLを取得
'【作成日】2023/02/09
Public Function GetAdminUpdateSQL() As String
On Error GoTo GetAdminUpdateSQL_Err
    
    Dim strSQL As String
    
    'ユーザーテーブルにデータを挿入するSQL
    strSQL = vbNullString
    strSQL = strSQL & "UPDATE "
    strSQL = strSQL & " Users "
    strSQL = strSQL & "SET"
    strSQL = strSQL & " Admin = ?"
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & " Id = ?"
    
    GetAdminUpdateSQL = strSQL
    
GetAdminUpdateSQL_Exit:
    Exit Function
GetAdminUpdateSQL_Err:
    GoTo GetAdminUpdateSQL_Exit
End Function


'【概要】ロック状態更新用SQLを取得
'【作成日】2023/02/18
Public Function GetLockUpdateSQL() As String
On Error GoTo GetLockUpdateSQL_Err
    
    Dim strSQL As String
    
    'ユーザーテーブルにデータを挿入するSQL
    strSQL = vbNullString
    strSQL = strSQL & "UPDATE "
    strSQL = strSQL & " Users "
    strSQL = strSQL & "SET"
    strSQL = strSQL & " LockFlag = ?"
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & " Id = ?"
    
    GetLockUpdateSQL = strSQL
    
GetLockUpdateSQL_Exit:
    Exit Function
GetLockUpdateSQL_Err:
    GoTo GetLockUpdateSQL_Exit
End Function


'【概要】 勤怠入力用SQLを取得
'【作成日】2023/02/18
Public Function GetAttendanceInsertSQL() As String
On Error GoTo GetAttendanceInsertSQL_Err
    
    Dim strSQL As String
    
    'ユーザーテーブルにデータを挿入するSQL
    strSQL = vbNullString
    strSQL = strSQL & "INSERT INTO Attendances "
    strSQL = strSQL & "("
    strSQL = strSQL & " EnvaringTime,"
    strSQL = strSQL & " Type, "
    strSQL = strSQL & " Apploval_Flag, "
    strSQL = strSQL & " UserId "
    strSQL = strSQL & ")"
    strSQL = strSQL & "VALUES "
    strSQL = strSQL & " (?, "
    strSQL = strSQL & "  ?, "
    strSQL = strSQL & "  ?, "
    strSQL = strSQL & "  ?) "
    
    GetAttendanceInsertSQL = strSQL
    
GetAttendanceInsertSQL_Exit:
    Exit Function
GetAttendanceInsertSQL_Err:
    GoTo GetAttendanceInsertSQL_Exit
End Function


'【概要】勤怠取得用SQLを取得
'【作成日】2023/02/21
Public Function GetAttendancesSQL() As String
On Error GoTo GetAttendancesSQL_Err
    
    Dim strSQL As String
    
    'ユーザーテーブルにデータを挿入するSQL
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " Id, "
    strSQL = strSQL & " EnvaringTime "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " Attendances "
    strSQL = strSQL & "ORDER BY "
    strSQL = strSQL & " Id ASC "
    
    GetAttendancesSQL = strSQL
    
GetAttendancesSQL_Exit:
    Exit Function
GetAttendancesSQL_Err:
    GoTo GetAttendancesSQL_Exit
End Function
