VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_SALES"
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
Private mvarCustomerID As Long 'local copy
Private mvarSalesNo As String 'local copy
Private mvarSalesDate As Date 'local copy
Private mvarTotal As Double 'local copy
Private mvarProductID As Long 'local copy
Private mvarQty As Long 'local copy
Private mvarSellingPrice As Double 'local copy
Private mvarTotalSales As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarUnitsSold As Long 'local copy
Private mvarSalesTime As Date 'local copy
'local variable(s) to hold property value(s)
Private mvarSalesDetailsID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarSellingUnit As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarStatus As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarTAX As Double 'local copy
Private mvarNHIL As Double 'local copy
Private mvarVAT As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarTotalQty As Long 'local copy
Public Property Let TotalQty(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalQty = 5
    mvarTotalQty = vData
End Property


Public Property Get TotalQty() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalQty
    TotalQty = mvarTotalQty
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



Public Property Let SellingUnit(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SellingUnit = 5
    mvarSellingUnit = vData
End Property


Public Property Get SellingUnit() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SellingUnit
    SellingUnit = mvarSellingUnit
End Property



Public Property Let SalesDetailsID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SalesDetailsID = 5
    mvarSalesDetailsID = vData
End Property


Public Property Get SalesDetailsID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SalesDetailsID
    SalesDetailsID = mvarSalesDetailsID
End Property



Public Property Let SalesTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SalesTime = 5
    mvarSalesTime = vData
End Property


Public Property Get SalesTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SalesTime
    SalesTime = mvarSalesTime
End Property



Public Property Let UnitsSold(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UnitsSold = 5
    mvarUnitsSold = vData
End Property


Public Property Get UnitsSold() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UnitsSold
    UnitsSold = mvarUnitsSold
End Property



Public Property Let TotalSales(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalSales = 5
    mvarTotalSales = vData
End Property


Public Property Get TotalSales() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalSales
    TotalSales = mvarTotalSales
End Property



Public Property Let SellingPrice(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UnitPrice = 5
    mvarSellingPrice = vData
End Property


Public Property Get SellingPrice() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UnitPrice
    SellingPrice = mvarSellingPrice
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



Public Property Let SalesDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SalesDate = 5
    mvarSalesDate = vData
End Property


Public Property Get SalesDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SalesDate
    SalesDate = mvarSalesDate
End Property



Public Property Let SalesNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SalesNo = 5
    mvarSalesNo = vData
End Property


Public Property Get SalesNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SalesNo
    SalesNo = mvarSalesNo
End Property



Public Property Let CustomerID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CustomerID = 5
    mvarCustomerID = vData
End Property


Public Property Get CustomerID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CustomerID
    CustomerID = mvarCustomerID
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

Public Function fn_LOAD_SALES(lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM tb_SALES"
        Else
            db.strQueryString = "SELECT * FROM tb_SALES WHERE SalesID = " & lngID
    End If

    Set rec = db.fn_MANAGE_RECORD

    Set fn_LOAD_SALES = rec

End Function


Public Function fn_LOAD_SALES_REPORT(Optional lngID As Long, Optional fromDate As Date, Optional toDate As Date) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM vw_SALES"
        Else
            db.strQueryString = "SELECT * FROM vw_SALES WHERE SalesDate BETWEEN '" & fromDate & "' AND '" & toDate & "'"
    End If

    Set rec = db.fn_MANAGE_RECORD

    Set fn_LOAD_SALES_REPORT = rec

End Function


Public Function fn_LOAD_PENDING_SALES(Optional lngID As Long) As ADODB.Recordset

    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM tb_SALES WHERE Status = 1"
        Else
            db.strQueryString = "SELECT * FROM vw_SALES WHERE  SalesID = " & lngID
    End If

    Set fn_LOAD_PENDING_SALES = db.fn_MANAGE_RECORD
    
End Function

Public Function fn_LOAD_SALES_DETAILS(lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM tb_SALES_DETAILS"
        Else
            db.strQueryString = "SELECT * FROM tb_ORDERS_DETAILS WHERE SalesID = " & lngID
    End If

    Set rec = db.fn_MANAGE_RECORD

    Set fn_LOAD_SALES_DETAILS = rec

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

Public Function fn_SAVE_SALES_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_SALES " _
    & "(SalesID,SalesNo,SalesDate,SalesTime,CustomerID,Status,TAX,Total) " _
    & " VALUES " _
    & "(" & mvarSalesID & "," _
    & "'" & mvarSalesNo & "'," _
    & "'" & mvarSalesDate & "'," _
    & "'" & mvarSalesTime & "'," _
    & mvarCustomerID & "," _
    & mvarStatus & "," _
    & mvarTAX & "," _
    & mvarTotalSales & ")"

    Set fn_SAVE_SALES_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_SAVE_SALES_DETAILS_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_SALES_DETAILS " _
    & "(SalesDetailsID,SalesID,ProductID,SellingUnit,Qty,TotalQty,SellingPrice,VAT,NHIL,Total) " _
    & " VALUES " _
    & "(" & fn_AUTOGEN_SALES_DETAILS_ID & "," _
    & mvarSalesID & "," _
    & mvarProductID & "," _
    & mvarSellingUnit & "," _
    & mvarQty & "," _
    & mvarTotalQty & "," _
    & mvarSellingPrice & "," _
    & mvarVAT & "," _
    & mvarNHIL & "," _
    & mvarTotal & ")"

    Set fn_SAVE_SALES_DETAILS_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_PRODUCTS_IN_STOCK(Optional lngProductID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_PRODUCTS SET " _
    & "UnitsInStock = UnitsInStock - " & mvarUnitsSold _
    & " WHERE ProductID = " & lngProductID
    
    Set fn_UPDATE_PRODUCTS_IN_STOCK = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_SALES_STATUS(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_SALES SET " _
    & "Status = 0 " _
    & " WHERE SalesID = " & lngID
    
    Set fn_UPDATE_SALES_STATUS = db.fn_MANAGE_RECORD

End Function

Public Function fn_AUTOGEN() As String

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(SalesNo) FROM tb_SALES"
    
    Set rec = db.fn_MANAGE_RECORD
    
    If rec.AbsolutePosition <> -1 Then
        If IsNull(rec(0)) Then
            fn_AUTOGEN = "S00001"
        Else
            fn_AUTOGEN = "S" & Format(Right(Trim(rec(0)), 5) + 1, "00000")
        End If
    Else
        fn_AUTOGEN = "S00001"
    End If

End Function

Public Function fn_ID_AUTOGEN() As Long

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(SalesID) FROM tb_SALES"
    
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

Public Function fn_AUTOGEN_SALES_DETAILS_ID() As Long

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(SalesDetailsID) FROM tb_SALES_DETAILS"
    
    Set rec = db.fn_MANAGE_RECORD
    
    If rec.AbsolutePosition <> -1 Then
        If IsNull(rec(0)) Then
            fn_AUTOGEN_SALES_DETAILS_ID = 1
        Else
            fn_AUTOGEN_SALES_DETAILS_ID = Val(rec(0)) + 1
        End If
    Else
        fn_AUTOGEN_SALES_DETAILS_ID = 1
    End If


End Function
