Option Compare Database


'【概要】文字数がおさまっているか
'【作成日】2023/02/02
Public Function IsFitInText(ByVal str As String, _
                        ByVal lngMaxLength As Long) As Boolean
On Error GoTo IsFitInText_Err

    '空の場合、False
    If str = vbNullString Then
        IsFitInText = False
        GoTo IsFitInText_Exit
    End If
    
    '制限文字数を超えている場合、False
    If Len(str) > lngMaxLength Then
        IsFitInText = False
        GoTo IsFitInText_Exit
    End If
    
    IsFitInText = True

IsFitInText_Exit:
    Exit Function
IsFitInText_Err:
    GoTo IsFitInText_Exit
End Function


'【概要】ログインしているか
'【作成日】2023/02/02
Public Function IsLogin() As Boolean
On Error GoTo IsLogin_Err

    IsLogin = False
    
    '初期値でない場合（設定されている場合）、True
    If lngLoginUserID <> 0 Then
        IsLogin = True
    End If

IsLogin_Exit:
    Exit Function
IsLogin_Err:
    GoTo IsLogin_Exit
End Function


'【概要】ログイン画面に飛ぶ（ログインしていない場合）
'【作成日】2023/02/02
'TODO :関数名改善の余地あり
Public Function LoginCheck(ByVal strOpenedFormName As String) As Boolean
On Error GoTo LoginCheck_Err

    LoginCheck = False
    
    'ログインしていない場合
    If IsLogin = False Then
        '自画面を非表示にする
        DoCmd.Close acForm, strOpenedFormName
        'ログイン画面を開く
        DoCmd.OpenForm LOGIN_FORM
        GoTo LoginCheck_Exit
    End If
        
    LoginCheck = True

LoginCheck_Exit:
    Exit Function
LoginCheck_Err:
    GoTo LoginCheck_Exit
End Function



'【概要】CSVファイル出力
'【作成日】2023/02/02
Public Function CSVOutput(ByVal strCSVOutputPath, _
                            ByVal strCSVFileName As String, _
                            ByVal strTblName As String) As Boolean
On Error GoTo CSVOutput_Err

    CSVOutput = False
    
    'CSV出力
    DoCmd.TransferText acExportDelim, , _
                        strTblName, _
                        strCSVOutputPath & strCSVFileName & " .csv", True
    
    CSVOutput = True

CSVOutput_Exit:
    Exit Function
CSVOutput_Err:
    GoTo CSVOutput_Exit
End Function
