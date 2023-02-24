Option Compare Database


'【概要】 ADOオブジェクト生成
Public Sub CreateCmdAndRs(ByRef objCmd As ADODB.Command, _
                        ByRef objRs As ADODB.Recordset)
On Error GoTo CreateCmdAndRs_Err
    
    'コマンド作成
    Call CreateCmd(objCmd)
    
    'レコードセット作成
    Call CreateRs(objRs)
    
CreateCmdAndRs_Err:

CreateCmdAndRs_Exit:
End Sub


'【概要】 コマンド作成
Public Sub CreateCmd(ByRef objCmd As ADODB.Command)
On Error GoTo CreateCmd_Err
    
    'コマンド
    Set objCmd = New ADODB.Command
    
CreateCmd_Err:

CreateCmd_Exit:
End Sub


'【概要】 レコードセット作成
Public Sub CreateRs(ByRef objRs As ADODB.Recordset)
On Error GoTo CreateRs_Err
    
    'レコードセット
    Set objRs = New ADODB.Recordset
    
CreateRs_Err:

CreateRs_Exit:
End Sub


'【概要】 コマンド破棄
Public Sub DisposeCmd(ByRef objCmd As ADODB.Command)
On Error GoTo DisposeCmd_Err
    
    '破棄
    Set objCmd = Nothing
    
DisposeCmd_Err:

DisposeCmd_Exit:
End Sub


'【概要】 レコードセット
Public Sub DisposeRs(ByRef objRs As ADODB.Recordset)
On Error GoTo DisposeRs_Err
    
    'レコードセットが存在する場合
    If Not objRs Is Nothing Then
        'レコードセットを閉じる
        objRs.Close
    End If
    
    '破棄
    Set objRs = Nothing
    
DisposeRs_Err:

DisposeRs_Exit:
End Sub



'【概要】 ADOオブジェクト破棄
Public Sub DisposeCmdAndRs(ByRef objCmd As ADODB.Command, _
                                ByRef objRs As ADODB.Recordset)
On Error GoTo DisposeCmdAndRs_Err
    
    'コマンド
    Call DisposeCmd(objCmd)
    
    'レコードセット
    Call DisposeRs(objRs)
    
DisposeCmdAndRs_Err:

DisposeCmdAndRs_Exit:
End Sub


'【概要】 ADOオブジェクトを生成
'TODO: 改善の余地あり
'==================================
Public Function ExecuteSQL(ByVal strSQL As String, _
                       ByRef objCmd As ADODB.Command, _
                       Optional ByRef objRs As ADODB.Recordset = Nothing) As Boolean
On Error GoTo ExecuteSQL_Err

    ExecuteSQL = False

    'コマンド
    With objCmd
        'コネクション
        .ActiveConnection = CurrentProject.Connection
        'コマンドテクスト
        .CommandText = strSQL
        If objRs Is Nothing Then
            '実行
            .Execute
        Else
            '実行結果を取得
            Set objRs = .Execute
        End If
    End With
    
    ExecuteSQL = True
    
ExecuteSQL_Exit:
    Exit Function
ExecuteSQL_Err:
    GoTo ExecuteSQL_Exit
End Function


'【概要】 パラメータを作成
'==================================
Public Sub AddParameter(ByRef objCmd As ADODB.Command, _
                    ByVal strParamName As String, _
                    ByVal strParamType As Long, _
                    ByVal strParamSize As Long, _
                    ByVal strParamValue As String)
On Error GoTo AddParameter_Err

    Dim objParam As ADODB.Parameter
    
    'パラメータ取得
    Set objParam = New ADODB.Parameter
    
    'パラメータを作成
    Set objParam = objCmd.CreateParameter(strParamName, strParamType, adParamInput, strParamSize, strParamValue)
    'パラメータ設定
    objCmd.Parameters.Append objParam
    
AddParameter_Exit:
    Set objParam = Nothing
    Exit Sub
AddParameter_Err:
    GoTo AddParameter_Exit
End Sub


'【概要】 データが存在するかどうか
'【作成日】2023/02/02
Public Function IsExistsData(ByVal objRs As ADODB.Recordset) As Boolean
On Error GoTo IsExistsData_Err
    
    IsExistsData = False
    
    'データが存在する場合、True
    If objRs.EOF = False Then
        IsExistsData = True
    End If
    
IsExistsData_Err:

IsExistsData_Exit:
End Function



