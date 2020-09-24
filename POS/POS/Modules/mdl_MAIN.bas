Attribute VB_Name = "mdl_MAIN"


Public Sub Main()

    Call sub_LOAD_SETTINGS
    Call fn_OPEN_CONNECTION

    frm_MAIN.Show
    frm_LOGIN.Show 1
    
End Sub

Sub sub_LOAD_SETTINGS()

    With SystemData
    
        .DB_ServerName = ReadIni(App.Title, "DB_ServerName")
        .DB_Database = ReadIni(App.Title, "DB_Database")
        .DB_UserName = ReadIni(App.Title, "DB_UserName")
'        .DB_Password = ReadIni(App.Title, "DB_Password")
    
    End With
    
End Sub


Public Function fn_OPEN_CONNECTION() As ADODB.Connection
On Error GoTo errHandler

    Dim con_Obj As New ADODB.Connection

    With SystemData
        con_Obj.ConnectionString = "PROVIDER=SQlOLEDB;Server=" & .DB_ServerName & ";UID=" & .DB_UserName & ";PWD=" & .DB_Password & ";Database=" & .DB_Database & ";"
    End With

    con_Obj.Open
    Set fn_OPEN_CONNECTION = con_Obj

    
EXITPROCEDURE:
    Exit Function
    
errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "ClsDatabase", "DBConnection")
    GoTo EXITPROCEDURE
    
End Function

Public Function fn_CLOSE_CONNECTION(ByRef con_Obj As ADODB.Connection)
On Error GoTo errHandler
    
    con_Obj.Close
    Set con_Obj = Nothing
    
EXITPROCEDURE:
    Exit Function
    
errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "ClsDatabase", "DBConnection")
    GoTo EXITPROCEDURE
    
End Function
