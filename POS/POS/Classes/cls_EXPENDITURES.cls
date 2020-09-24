VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_EXPENDITURES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarExpenditureID As Long 'local copy
Private mvarExpenditureTypeID As Long 'local copy
Private mvarExpenditureDate As Date 'local copy
Private mvarExpenditureTime As Date 'local copy
Private mvarExpenditureTotal As Double 'local copy
Private mvarExpenditureQty As Variant 'local copy
Private mvarExpenditureItems As String 'local copy
Private mvarExpenditurePrice As Double 'local copy
Private mvarExpenditureDescription As String 'local copy
Private mvarExpenditureDetailsTotal As Double 'local copy
Public Property Let ExpenditureDetailsTotal(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExpenditureDetailsTotal = 5
    mvarExpenditureDetailsTotal = vData
End Property


Public Property Get ExpenditureDetailsTotal() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExpenditureDetailsTotal
    ExpenditureDetailsTotal = mvarExpenditureDetailsTotal
End Property



Public Property Let ExpenditureDescription(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExpenditureDescription = 5
    mvarExpenditureDescription = vData
End Property


Public Property Get ExpenditureDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExpenditureDescription
    ExpenditureDescription = mvarExpenditureDescription
End Property



Public Property Let ExpenditurePrice(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExpenditurePrice = 5
    mvarExpenditurePrice = vData
End Property


Public Property Get ExpenditurePrice() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExpenditurePrice
    ExpenditurePrice = mvarExpenditurePrice
End Property



Public Property Let ExpenditureItems(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExpenditureItems = 5
    mvarExpenditureItems = vData
End Property


Public Property Get ExpenditureItems() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExpenditureItems
    ExpenditureItems = mvarExpenditureItems
End Property



Public Property Let ExpenditureQty(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExpenditureQty = 5
    mvarExpenditureQty = vData
End Property


Public Property Set ExpenditureQty(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ExpenditureQty = Form1
    Set mvarExpenditureQty = vData
End Property


Public Property Get ExpenditureQty() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExpenditureQty
    If IsObject(mvarExpenditureQty) Then
        Set ExpenditureQty = mvarExpenditureQty
    Else
        ExpenditureQty = mvarExpenditureQty
    End If
End Property



Public Property Let ExpenditureTotal(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExpenditureTotal = 5
    mvarExpenditureTotal = vData
End Property


Public Property Get ExpenditureTotal() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExpenditureTotal
    ExpenditureTotal = mvarExpenditureTotal
End Property



Public Property Let ExpenditureTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExpenditureTime = 5
    mvarExpenditureTime = vData
End Property


Public Property Get ExpenditureTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExpenditureTime
    ExpenditureTime = mvarExpenditureTime
End Property



Public Property Let ExpenditureDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExpenditureDate = 5
    mvarExpenditureDate = vData
End Property


Public Property Get ExpenditureDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExpenditureDate
    ExpenditureDate = mvarExpenditureDate
End Property



Public Property Let ExpenditureTypeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExpenditureTypeID = 5
    mvarExpenditureTypeID = vData
End Property


Public Property Get ExpenditureTypeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExpenditureTypeID
    ExpenditureTypeID = mvarExpenditureTypeID
End Property



Public Property Let ExpenditureID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExpenditureID = 5
    mvarExpenditureID = vData
End Property


Public Property Get ExpenditureID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExpenditureID
    ExpenditureID = mvarExpenditureID
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

Public Function fn_SAVE_EXPENDITURES_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_EXPENDITURES " _
    & "(ExpenditureID,ExpenditureTypeID,ExpenditureTotal,ExpenditureDate,ExpenditureTime) " _
    & " VALUES " _
    & "(" & mvarExpenditureID & "," _
    & mvarExpenditureTypeID & "," _
    & mvarExpenditureTotal & "," _
    & "'" & mvarExpenditureDate & "'," _
    & "'" & mvarExpenditureTime & "')"

    Set fn_SAVE_EXPENDITURES_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_SAVE_EXPENDITURES_DETAILS_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_EXPENDITURE_DETAILS " _
    & "(ExpenditureID,ExpenditureItems,ExpenditureQty,ExpenditurePrice,ExpenditureDescription,ExpenditureTotal) " _
    & " VALUES " _
    & "(" & mvarExpenditureID & "," _
    & "'" & mvarExpenditureItems & "'," _
    & mvarExpenditureQty & "," _
    & mvarExpenditurePrice & "," _
    & "'" & mvarExpenditureDescription & "'," _
    & mvarExpenditureDetailsTotal & ")"

    Set fn_SAVE_EXPENDITURES_DETAILS_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fN_LOAD_EXPENDITURES_REPORT(Optional lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset

    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM vw_EXPENDITURES "
        Else
            db.strQueryString = "SELECT * FROM vw_EXPENDITURES WHERE ExpenditureID = " & lngID
    End If
    
    Set fN_LOAD_EXPENDITURES_REPORT = db.fn_MANAGE_RECORD


End Function


Public Function fn_AUTOGEN() As Long

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(ExpenditureID) FROM tb_EXPENDITURES"
    
    Set rec = db.fn_MANAGE_RECORD
    
    If rec.AbsolutePosition <> -1 Then
        If IsNull(rec(0)) Then
            fn_AUTOGEN = 1
        Else
            fn_AUTOGEN = Val(rec(0)) + 1
        End If
    Else
        fn_AUTOGEN = 1
    End If


End Function

