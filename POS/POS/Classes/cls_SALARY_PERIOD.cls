VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_SALARY_PERIOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarStartDate As Date 'local copy
Private mvarEndDate As Date 'local copy
'local variable(s) to hold property value(s)
Private mvarSalaryPeriodDetailsID As Long 'local copy
Private mvarSalaryID As Long 'local copy
Private mvarSalaryPeriodID As Long 'local copy
Public Property Let SalaryPeriodID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SalaryPeriodID = 5
    mvarSalaryPeriodID = vData
End Property


Public Property Get SalaryPeriodID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SalaryPeriodID
    SalaryPeriodID = mvarSalaryPeriodID
End Property



Public Property Let SalaryID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SalaryID = 5
    mvarSalaryID = vData
End Property


Public Property Get SalaryID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SalaryID
    SalaryID = mvarSalaryID
End Property



Public Property Let SalaryPeriodDetailsID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SalaryPeriodDetailsID = 5
    mvarSalaryPeriodDetailsID = vData
End Property


Public Property Get SalaryPeriodDetailsID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SalaryPeriodDetailsID
    SalaryPeriodDetailsID = mvarSalaryPeriodDetailsID
End Property



Public Property Let EndDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EndDate = 5
    mvarEndDate = vData
End Property


Public Property Get EndDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EndDate
    EndDate = mvarEndDate
End Property



Public Property Let StartDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartDate = 5
    mvarStartDate = vData
End Property


Public Property Get StartDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartDate
    StartDate = mvarStartDate
End Property


Private Sub Class_Initialize()
'On Error GoTo errHandler

    Set con_Obj = mdl_MAIN.fn_OPEN_CONNECTION

EXITPROCEDURE:
    Exit Sub
    
'errHandler:
'    MsgBox Err.Description, vbCritical, "Connection"
'    Call MdlFunctions.fnWriteErrorToFile(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
'    GoTo EXITPROCEDURE
End Sub

Private Sub Class_Terminate()
On Error GoTo errHandler

    Call mdl_MAIN.fn_CLOSE_CONNECTION(con_Obj)

EXITPROCEDURE:
    Exit Sub
    
errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub

Public Function fn_SAVE_SALARY_PERIOD() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_SALARY_PERIOD " _
    & "(StartDate,EndDate) " _
    & " VALUES " _
    & "('" & mvarStartDate & "'," _
    & "'" & mvarEndDate & "')"

    Set fn_SAVE_SALARY_PERIOD = db.fn_MANAGE_RECORD

End Function



Public Function fn_UPDATE_SALARY_PERIOD(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_SALARY_PERIOD SET " _
    & "StartDate = '" & mvarStartDate & "'," _
    & "EndDate = '" & mvarEndDate & "'" _
    & " WHERE SalaryPeriodID = " & lngID
    
    Set fn_UPDATE_SALARY_PERIOD = db.fn_MANAGE_RECORD

End Function

Public Function fN_LOAD_SALARY_PERIOD(Optional lngID As Long) As ADODB.Recordset

    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM tb_SALARY_PERIOD "
        Else
            db.strQueryString = "SELECT * FROM tb_SALARY_PERIOD WHERE SalaryPeriodID = " & lngID
    End If
    
    Set fN_LOAD_SALARY_PERIOD = db.fn_MANAGE_RECORD

End Function


Public Function fn_DELETE_SALARY_PERIOD(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = "DELETE FROM tb_SALARY_PERIOD " _
    & " WHERE SalaryPeriodID = " & lngID
    
    Set fn_DELETE_SALARY_PERIOD = db.fn_MANAGE_RECORD

End Function
