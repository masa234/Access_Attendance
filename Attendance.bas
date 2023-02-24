Option Compare Database


'【概要】勤怠入力
'【作成日】2023/02/02
Public Function AttendanceInsert(ByVal lngUserID As String, _
                                ByVal lngType As Long) As Boolean
On Error GoTo AttendanceInsert_Err

    AttendanceInsert = False
    
    Dim objCmd As ADODB.Command
    Dim strSQL As String
    Dim strCurrentDate As String
    
    'オブジェクト作成
    Call CreateCmd(objCmd)
    
    'SQL取得
    strSQL = GetAttendanceInsertSQL
    
    '現在時刻
    'TODO ;　共通化した方がよいかも
    strCurrentDateTime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    
    'パラム設定
    Call AddParameter(objCmd, "打刻時間", adChar, 255, strCurrentDateTime)
    Call AddParameter(objCmd, "勤怠種別", adChar, 255, lngType)
    Call AddParameter(objCmd, "承認フラグ", adChar, 255, EnumIsApploval.NotApploval)
    Call AddParameter(objCmd, "ユーザID", adChar, 255, lngUserID)
    
    'SQL実行
    'TODO :動作しない
    If ExecuteSQL(strSQL, objCmd) = False Then
        GoTo AttendanceInsert_Exit
    End If
    
AttendanceInsert_Exit:
    'オブジェクトを破棄
    Call DisposeCmd(objCmd)
    Exit Function
AttendanceInsert_Err:
    GoTo AttendanceInsert_Exit
End Function

