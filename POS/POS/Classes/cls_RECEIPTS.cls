VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_RECEIPTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarSalesID As Long 'local copy
Private mvarGrossAmount As Double 'local copy
Private mvarDiscount As Double 'local copy
Private mvarNetAmount As Double 'local copy
Private mvarAmountPaid As Double 'local copy
Private mvarBalance As Double 'local copy
Public Property Let Balance(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Balance = 5
    mvarBalance = vData
End Property


Public Property Get Balance() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Balance
    Balance = mvarBalance
End Property



Public Property Let AmountPaid(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AmountPaid = 5
    mvarAmountPaid = vData
End Property


Public Property Get AmountPaid() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AmountPaid
    AmountPaid = mvarAmountPaid
End Property



Public Property Let NetAmount(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NetAmount = 5
    mvarNetAmount = vData
End Property


Public Property Get NetAmount() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NetAmount
    NetAmount = mvarNetAmount
End Property



Public Property Let Discount(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Discount = 5
    mvarDiscount = vData
End Property


Public Property Get Discount() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Discount
    Discount = mvarDiscount
End Property



Public Property Let GrossAmount(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GrossAmount = 5
    mvarGrossAmount = vData
End Property


Public Property Get GrossAmount() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GrossAmount
    GrossAmount = mvarGrossAmount
End Property



Public Property Let SalesID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SalesID = 5
    mvarSalesID = vData
End Property


Public Property Get SalesID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SalesID
    SalesID = mvarSalesID
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

Public Function fn_SAVE_RECEIPT_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_RECEIPTS " _
    & "(SalesID,GrossAmount,Discount,NetAmount,AmountPaid,Balance) " _
    & " VALUES " _
    & "(" & mvarSalesID & "," _
    & mvarGrossAmount & "," _
    & mvarDiscount & "," _
    & mvarNetAmount & "," _
    & mvarAmountPaid & "," _
    & mvarBalance & ")"

    Set fn_SAVE_RECEIPT_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_LOAD_RECEIPTS_DETAILS(lngID As Long) As ADODB.Recordset

    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM vw_SALES_RECEIPTS"
        Else
            db.strQueryString = "SELECT * FROM vw_SALES_RECEIPTS WHERE SalesID = " & lngID
    End If

    Set fn_LOAD_RECEIPTS_DETAILS = db.fn_MANAGE_RECORD


End Function
