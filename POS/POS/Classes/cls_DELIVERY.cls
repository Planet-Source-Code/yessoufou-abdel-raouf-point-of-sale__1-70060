VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_DELIVERY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Dim con_Obj As New ADODB.Connection
Private mvarDeliveryID As Long 'local copy
Private mvarOrderID As Long 'local copy
Private mvarDeliveryDate As Date 'local copy
Private mvarProductID As Long 'local copy
Private mvarUnitPrice As Double 'local copy
Private mvarUnitsOrdered As Long 'local copy
Private mvarUnitsReceived As Double 'local copy
Private mvarDeliveryTime As Date 'local copy
'local variable(s) to hold property value(s)
Private mvarSupplierPrice As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarVAT As Double 'local copy
Private mvarNHIL As Double 'local copy
Private mvarTAX As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarTotalWithoutTax As Double 'local copy
Private mvarTotalWithTax As Double 'local copy
Public Property Let TotalWithTax(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalWithTax = 5
    mvarTotalWithTax = vData
End Property


Public Property Get TotalWithTax() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalWithTax
    TotalWithTax = mvarTotalWithTax
End Property



Public Property Let TotalWithoutTax(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalWithoutTax = 5
    mvarTotalWithoutTax = vData
End Property


Public Property Get TotalWithoutTax() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalWithoutTax
    TotalWithoutTax = mvarTotalWithoutTax
End Property








Public Property Let TAX(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TAX = 5
    mvarTAX = vData
End Property


Public Property Get TAX() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TAX
    TAX = mvarTAX
End Property



Public Property Let NHIL(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NHIL = 5
    mvarNHIL = vData
End Property


Public Property Get NHIL() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NHIL
    NHIL = mvarNHIL
End Property



Public Property Let VAT(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VAT = 5
    mvarVAT = vData
End Property


Public Property Get VAT() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VAT
    VAT = mvarVAT
End Property



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



Public Property Let DeliveryTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DeliveryTime = 5
    mvarDeliveryTime = vData
End Property


Public Property Get DeliveryTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DeliveryTime
    DeliveryTime = mvarDeliveryTime
End Property



Public Property Let UnitsReceived(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UnitsReceived = 5
    mvarUnitsReceived = vData
End Property


Public Property Get UnitsReceived() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UnitsReceived
    UnitsReceived = mvarUnitsReceived
End Property



Public Property Let UnitsOrdered(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UnitsOrdered = 5
    mvarUnitsOrdered = vData
End Property


Public Property Get UnitsOrdered() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UnitsOrdered
    UnitsOrdered = mvarUnitsOrdered
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



Public Property Let DeliveryDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DeliveryDate = 5
    mvarDeliveryDate = vData
End Property


Public Property Get DeliveryDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DeliveryDate
    DeliveryDate = mvarDeliveryDate
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



Public Property Let DeliveryID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DeliveryID = 5
    mvarDeliveryID = vData
End Property


Public Property Get DeliveryID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DeliveryID
    DeliveryID = mvarDeliveryID
End Property


Public Function fn_LOAD_ORDERS(lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM tb_ORDERS WHERE Status = 0"
        Else
            db.strQueryString = "SELECT * FROM tb_ORDERS WHERE Status = 0 AND OrderID = " & lngID
    End If

    Set rec = db.fn_MANAGE_RECORD

    Set fn_LOAD_ORDERS = rec

End Function

Public Function fn_LOAD_DELIVERY_REPORT(lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM vw_DELIVERY "
        Else
'            db.strQueryString = "SELECT * FROM vw_DELIVERY WHERE Status = 0 AND OrderID = " & lngID
    End If

    Set rec = db.fn_MANAGE_RECORD

    Set fn_LOAD_DELIVERY_REPORT = rec

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

Public Function fn_SAVE_DELIVERY_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_DELIVERY " _
    & "(DeliveryID,OrderID,DeliveryDate,TotalWithoutTax,Tax,TotalWithTax,DeliveryTime) " _
    & " VALUES " _
    & "(" & mvarDeliveryID & "," _
    & mvarOrderID & "," _
    & "'" & mvarDeliveryDate & "'," _
    & mvarTotalWithoutTax & "," _
    & mvarTAX & "," _
    & mvarTotalWithTax & "," _
    & "'" & mvarDeliveryTime & "')"

    Set fn_SAVE_DELIVERY_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_SAVE_DELIVERY_DETAILS_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_DELIVERY_DETAILS " _
    & "(DeliveryDetailsID,DeliveryID,ProductID,supplierPrice,UnitsOrdered,UnitsReceived,VAT,NHIL) " _
    & " VALUES " _
    & "(" & fn_AUTOGEN_DELIVERY_DETAILS_ID & "," _
    & mvarDeliveryID & "," _
    & mvarProductID & "," _
    & mvarSupplierPrice & "," _
    & mvarUnitsOrdered & "," _
    & mvarUnitsReceived & "," _
    & mvarVAT & "," _
    & mvarNHIL & ")"

    Set fn_SAVE_DELIVERY_DETAILS_RECORDS = db.fn_MANAGE_RECORD

End Function


Public Function fn_AUTOGEN() As Long

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(DeliveryID) FROM tb_DELIVERY"
    
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

Public Function fn_AUTOGEN_DELIVERY_DETAILS_ID() As Long

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(DeliveryDetailsID) FROM tb_DELIVERY_DETAILS"
    
    Set rec = db.fn_MANAGE_RECORD
    
    If rec.AbsolutePosition <> -1 Then
        If IsNull(rec(0)) Then
            fn_AUTOGEN_DELIVERY_DETAILS_ID = 1
        Else
            fn_AUTOGEN_DELIVERY_DETAILS_ID = Val(rec(0)) + 1
        End If
    Else
        fn_AUTOGEN_DELIVERY_DETAILS_ID = 1
    End If


End Function

Public Function fn_UPDATE_PRODUCTS_IN_STOCK(Optional lngProductID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_PRODUCTS SET " _
    & "UnitsInStock = UnitsInStock + " & mvarUnitsReceived _
    & " WHERE ProductID = " & lngProductID
    
    Set fn_UPDATE_PRODUCTS_IN_STOCK = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_ORDERS_STATUS(Optional lngOrderID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_ORDERS SET " _
    & "Status =1" _
    & " WHERE OrderID = " & lngOrderID
    
    Set fn_UPDATE_ORDERS_STATUS = db.fn_MANAGE_RECORD

End Function

Public Function fn_LOAD_DELIVERY_TAX_REPORT(Optional lngID As Long, Optional fromDate As Date, Optional toDate As Date) As ADODB.Recordset

    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM vw_DELIVERY"
        Else
            db.strQueryString = "SELECT * FROM vw_DELIVERY WHERE DeliveryDate BETWEEN '" & fromDate & "' AND '" & toDate & "'"
    End If

    Set fn_LOAD_DELIVERY_TAX_REPORT = db.fn_MANAGE_RECORD

End Function
