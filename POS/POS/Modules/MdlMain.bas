Attribute VB_Name = "MdlMain"


Public Sub Main()

    Call subLoadRegistrySettings
    Call fnOpenConnection

    frm_MAIN.Show
    frm_LOGIN.Show 1
    
End Sub

'**********Function to load registry details***********
Public Sub subLoadRegistrySettings()
On Error GoTo errHandler

    With registrySettings
        .Server = GetSetting("AbdelSoft", "BMS", "Server", "")
        .Dababase = GetSetting("AbdelSoft", "BMS", "Database", "")
        .UserName = GetSetting("AbdelSoft", "BMS", "UserName", "")
        .Password = GetSetting("AbdelSoft", "BMS", "Password", "")
    End With

EXITPROCEDURE:
    Exit Sub
    
errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call MdlFunctions.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub




Public Function fnOpenConnection() As ADODB.Connection
On Error GoTo errHandler

    Dim con_Obj As New ADODB.Connection

    With registrySettings
        con_Obj.ConnectionString = "provider = microsoft.jet.oledb.4.0 ; data source = " & App.Path & "\Database\" & .Dababase & ".mdb;Jet OLEDB:Database Password=" & .Password
    End With
    
    con_Obj.Open
    Set fnOpenConnection = con_Obj

    
EXITPROCEDURE:
    Exit Function
    
errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call MdlFunctions.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "ClsDatabase", "DBConnection")
    GoTo EXITPROCEDURE
    
End Function

Public Function fnCloseConnection(ByRef con_Obj As ADODB.Connection)
On Error GoTo errHandler
    
    con_Obj.Close
    Set con_Obj = Nothing
    
EXITPROCEDURE:
    Exit Function
    
errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call MdlFunctions.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "ClsDatabase", "DBConnection")
    GoTo EXITPROCEDURE
    
End Function
