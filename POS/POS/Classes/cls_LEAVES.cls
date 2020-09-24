VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_LEAVES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarEmployeeID As Long 'local copy
Private mvarEmployeeNo As String 'local copy
Private mvarLeaveID As Long 'local copy
Private mvarLeaveName As String 'local copy
Private mvarApprovedDate As Date 'local copy
Private mvarStartDate As Date 'local copy
Private mvarEndDate As Date 'local copy
Private mvarNotes As String 'local copy
Public Property Let Notes(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Notes = 5
    mvarNotes = vData
End Property


Public Property Get Notes() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Notes
     Notes = mvarNotes
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



Public Property Let ApprovedDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ApprovedDate = 5
    mvarApprovedDate = vData
End Property


Public Property Get ApprovedDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ApprovedDate
    ApprovedDate = mvarApprovedDate
End Property



Public Property Let LeaveName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LeaveName = 5
    mvarLeaveName = vData
End Property


Public Property Get LeaveName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LeaveName
    LeaveName = mvarLeaveName
End Property



Public Property Let LeaveID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LeaveID = 5
    mvarLeaveID = vData
End Property


Public Property Get LeaveID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LeaveID
    LeaveID = mvarLeaveID
End Property



Public Property Let EmployeeNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EmployeeNo = 5
    mvarEmployeeNo = vData
End Property


Public Property Get EmployeeNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EmployeeNo
    EmployeeNo = mvarEmployeeNo
End Property



Public Property Let EmployeeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EmployeeID = 5
    mvarEmployeeID = vData
End Property


Public Property Get EmployeeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EmployeeID
    EmployeeID = mvarEmployeeID
End Property

Private Sub Class_Initialize()
'On Error GoTo errHandler

    Set con_Obj = mdl_MAIN.fn_OPEN_CONNECTION

EXITPROCEDURE:
    Exit Sub
    
'errHandler:
'    MsgBox Err.Description, vbCritical, "Connection"
'    Call Mdl_Functions.fnWriteErrorToFile(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
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

Public Function fn_LOAD_LEAVES(Optional lngEmployeeID As Long) As ADODB.Recordset
        
    db.strQueryString = "SELECT * FROM vw_LEAVES WHERE EmployeeID = " & lngEmployeeID
    
    Set fn_LOAD_LEAVES = db.fn_MANAGE_RECORD
    
End Function

Public Function fn_DELETE_LEAVES(Optional lngEmployeeID As Long, Optional lngLeaveDetailsID As Long) As ADODB.Recordset
        
    db.strQueryString = "DELETE FROM tb_LEAVES_DETAILS WHERE EmployeeID = " & lngEmployeeID & " AND LeaveDetailsID = " & lngLeaveDetailsID
    
    Set fn_DELETE_LEAVES = db.fn_MANAGE_RECORD
    
End Function

Public Function fn_LOAD_LEAVES_DETAILS(Optional lngEmployeeID As Long, Optional lngLeaveDetailsID As Long) As ADODB.Recordset
        
    db.strQueryString = "SELECT * FROM vw_LEAVES WHERE EmployeeID = " & lngEmployeeID & " AND LeaveDetailsID = " & lngLeaveDetailsID
    
    Set fn_LOAD_LEAVES_DETAILS = db.fn_MANAGE_RECORD
    
End Function

Public Function fn_SAVE_LEAVES_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_LEAVES_DETAILS " _
    & "(LeaveID,ApprovedDate,StartDate,EndDate,Notes,EmployeeID) " _
    & " VALUES " _
    & "(" & mvarLeaveID & "," _
    & "'" & mvarApprovedDate & "'," _
    & "'" & mvarStartDate & "'," _
    & "'" & mvarEndDate & "'," _
    & "'" & mvarNotes & "'," _
    & mvarEmployeeID & ")"

    Set fn_SAVE_LEAVES_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_LEAVES_RECORDS(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_LEAVES_DETAILS SET " _
    & "LeaveID = " & mvarLeaveID & "," _
    & "ApprovedDate = '" & mvarApprovedDate & "'," _
    & "StartDate = '" & mvarStartDate & "'," _
    & "EndDate = '" & mvarEndDate & "'," _
    & "Notes = '" & mvarNotes & "'," _
    & "EmployeeID = " & mvarEmployeeID _
    & " WHERE LeaveDetailsID = " & lngID
    
    Set fn_UPDATE_LEAVES_RECORDS = db.fn_MANAGE_RECORD

End Function

