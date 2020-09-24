VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_ORDERS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarOrderID As Long 'local copy
Private mvarOrderNo As String 'local copy
Private mvarOrderDate As Date 'local copy
Private mvarSupplierID As Long 'local copy
Private mvarProductID As Long 'local copy
Private mvarQty As Long 'local copy
Private mvarUnitPrice As Double 'local copy
Private mvarTotal As Double 'local copy
Private mvarTotalOrder As Double 'local copy
Private mvarStatus As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarOrderTime As Date 'local copy
'local variable(s) to hold property value(s)
Private mvarBuyingUnit As Long 'local copy
Private mvarSupplierPrice As Double 'local copy
Public Property Let SupplierPrice(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SupplierPrice = 5
    mvarSupplierPrice = vData
End Property


Public Property Get SupplierPrice() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SupplierPrice
    SupplierPrice = mvarSupplierPrice
End Property



Public Property Let BuyingUnit(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BuyingUnit = 5
    mvarBuyingUnit = vData
End Property


Public Property Get BuyingUnit() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuyingUnit
    BuyingUnit = mvarBuyingUnit
End Property



Public Property Let OrderTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OrderTime = 5
    mvarOrderTime = vData
End Property


Public Property Get OrderTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OrderTime
    OrderTime = mvarOrderTime
End Property



Public Property Let Status(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Status = 5
    mvarStatus = vData
End Property


Public Property Get Status() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Status
    Status = mvarStatus
End Property



Public Property Let TotalOrder(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalOrder = 5
    mvarTotalOrder = vData
End Property


Public Property Get TotalOrder() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalOrder
    TotalOrder = mvarTotalOrder
End Property



Public Property Let Total(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Total = 5
    mvarTotal = vData
End Property


Public Property Get Total() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Total
    Total = mvarTotal
End Property



Public Property Let UnitPrice(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UnitPrice = 5
    mvarUnitPrice = vData
End Property


Public Property Get UnitPrice() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UnitPrice
    UnitPrice = mvarUnitPrice
End Property



Public Property Let Qty(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Qty = 5
    mvarQty = vData
End Property


Public Property Get Qty() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Qty
    Qty = mvarQty
End Property



Public Property Let ProductID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProductID = 5
    mvarProductID = vData
End Property


Public Property Get ProductID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductID
    ProductID = mvarProductID
End Property



Public Property Let SupplierID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SupplierID = 5
    mvarSupplierID = vData
End Property


Public Property Get SupplierID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SupplierID
    SupplierID = mvarSupplierID
End Property



Public Property Let OrderDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OrderDate = 5
    mvarOrderDate = vData
End Property


Public Property Get OrderDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OrderDate
    OrderDate = mvarOrderDate
End Property



Public Property Let OrderNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OrderNo = 5
    mvarOrderNo = vData
End Property


Public Property Get OrderNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OrderNo
    OrderNo = mvarOrderNo
End Property



Public Property Let OrderID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OrderID = 5
    mvarOrderID = vData
End Property


Public Property Get OrderID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OrderID
    OrderID = mvarOrderID
End Property

Public Function fn_LOAD_ORDERS(lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM tb_ORDERS"
        Else
            db.strQueryString = "SELECT * FROM tb_ORDERS WHERE OrderID = " & lngID
    End If

    Set rec = db.fn_MANAGE_RECORD

    Set fn_LOAD_ORDERS = rec

End Function

Public Function fn_LOAD_ORDERS_REPORT(lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM vw_ORDERS"
        Else
            db.strQueryString = "SELECT * FROM vw_ORDERS WHERE Status = " & lngID
    End If

    Set rec = db.fn_MANAGE_RECORD

    Set fn_LOAD_ORDERS_REPORT = rec

End Function

Public Function fn_LOAD_PENDING_ORDERS_REPORT() As ADODB.Recordset

    Dim rec As New ADODB.Recordset

    db.strQueryString = "SELECT * FROM vw_ORDERS WHERE Status = 0"

    Set rec = db.fn_MANAGE_RECORD

    Set fn_LOAD_PENDING_ORDERS_REPORT = rec

End Function

Public Function fn_LOAD_ORDERS_DETAILS(lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM tb_ORDERS_DETAILS"
        Else
            db.strQueryString = "SELECT * FROM tb_ORDERS_DETAILS WHERE OrderID = " & lngID
    End If

    Set rec = db.fn_MANAGE_RECORD

    Set fn_LOAD_ORDERS_DETAILS = rec

End Function

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

Public Function fn_SAVE_ORDERS_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_ORDERS " _
    & "(OrderID,OrderNo,OrderDate,OrderTime,SupplierID,status,Total) " _
    & " VALUES " _
    & "(" & mvarOrderID & "," _
    & "'" & mvarOrderNo & "'," _
    & "'" & mvarOrderDate & "'," _
    & "'" & mvarOrderTime & "'," _
    & mvarSupplierID & "," _
    & mvarStatus & "," _
    & mvarTotalOrder & ")"

    Set fn_SAVE_ORDERS_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_SAVE_ORDERS_DETAILS_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_ORDERS_DETAILS " _
    & "(OrderDetailsID,OrderID,ProductID,BuyingUnit,Qty,SupplierPrice,Total) " _
    & " VALUES " _
    & "(" & fn_AUTOGEN_ORDERS_DETAILS_ID & "," _
    & mvarOrderID & "," _
    & mvarProductID & "," _
    & mvarBuyingUnit & "," _
    & mvarQty & "," _
    & mvarSupplierPrice & "," _
    & mvarTotal & ")"

    Set fn_SAVE_ORDERS_DETAILS_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_AUTOGEN() As String

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(OrderNo) FROM tb_ORDERS"
    
    Set rec = db.fn_MANAGE_RECORD
    
    If rec.AbsolutePosition <> -1 Then
        If IsNull(rec(0)) Then
            fn_AUTOGEN = "O00001"
        Else
            fn_AUTOGEN = "O" & Format(Right(Trim(rec(0)), 5) + 1, "00000")
        End If
    Else
        fn_AUTOGEN = "O00001"
    End If

End Function

Public Function fn_ID_AUTOGEN() As Long

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(OrderID) FROM tb_ORDERS"
    
    Set rec = db.fn_MANAGE_RECORD
    
    If rec.AbsolutePosition <> -1 Then
        If IsNull(rec(0)) Then
            fn_ID_AUTOGEN = 1
        Else
            fn_ID_AUTOGEN = Val(rec(0)) + 1
        End If
    Else
        fn_ID_AUTOGEN = 1
    End If


End Function

Public Function fn_AUTOGEN_ORDERS_DETAILS_ID() As Long

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(OrderDetailsID) FROM tb_ORDERS_DETAILS"
    
    Set rec = db.fn_MANAGE_RECORD
    
    If rec.AbsolutePosition <> -1 Then
        If IsNull(rec(0)) Then
            fn_AUTOGEN_ORDERS_DETAILS_ID = 1
        Else
            fn_AUTOGEN_ORDERS_DETAILS_ID = Val(rec(0)) + 1
        End If
    Else
        fn_AUTOGEN_ORDERS_DETAILS_ID = 1
    End If


End Function
