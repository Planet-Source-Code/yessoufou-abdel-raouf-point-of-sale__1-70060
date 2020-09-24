VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_PRODUCT_PACKAGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarProductID As Long 'local copy
Private mvarPackageID As Long 'local copy
Private mvarSupplierPrice As Double 'local copy
Private mvarQty As Long 'local copy
Private mvarSellingPriceWithoutTax As Double 'local copy
Private mvarVAT As Double 'local copy
Private mvarNHIL As Double 'local copy
Private mvarSellingPriceWithTax As Double 'local copy
Public Property Let SellingPriceWithTax(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SellingPriceWithTax = 5
    mvarSellingPriceWithTax = vData
End Property


Public Property Get SellingPriceWithTax() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SellingPriceWithTax
    SellingPriceWithTax = mvarSellingPriceWithTax
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



Public Property Let SellingPriceWithoutTax(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SellingPriceWithoutTax = 5
    mvarSellingPriceWithoutTax = vData
End Property


Public Property Get SellingPriceWithoutTax() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SellingPriceWithoutTax
    SellingPriceWithoutTax = mvarSellingPriceWithoutTax
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



Public Property Let PackageID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PackageID = 5
    mvarPackageID = vData
End Property


Public Property Get PackageID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PackageID
    PackageID = mvarPackageID
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


Public Function fn_LOAD_PRODUCT_PACKAGE(Optional lngID As Long) As ADODB.Recordset

    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM tb_PRODUCT_PACKAGE "
        Else
            db.strQueryString = "SELECT * FROM tb_PRODUCT_PACKAGE WHERE ProductID = " & lngID
    End If
    
    Set fn_LOAD_PRODUCT_PACKAGE = db.fn_MANAGE_RECORD

End Function

Public Function fn_LOAD_PRODUCT_PACKAGE_DETAILS(Optional lngProductID As Long, Optional lngPackageID As Long) As ADODB.Recordset

    If lngPackageID = 0 Then
        db.strQueryString = "SELECT * FROM tb_PRODUCT_PACKAGE "
        Else
            db.strQueryString = "SELECT * FROM tb_PRODUCT_PACKAGE WHERE ProductID = " & lngProductID & " AND PackageID = " & lngPackageID
    End If
    
    Set fn_LOAD_PRODUCT_PACKAGE_DETAILS = db.fn_MANAGE_RECORD

End Function

Public Function fn_SAVE_PRODUCTS_PACKAGE() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_PRODUCT_PACKAGE " _
    & "(ProductID,PackageID,SupplierPrice,Qty,SellingPriceWithoutTax,VAT,NHIL,SellingPriceWithTax) " _
    & " VALUES " _
    & "(" & mvarProductID & "," _
    & mvarPackageID & "," _
    & mvarSupplierPrice & "," _
    & mvarQty & "," _
    & mvarSellingPriceWithoutTax & "," _
    & mvarVAT & "," _
    & mvarNHIL & "," _
    & mvarSellingPriceWithTax & ")"

    Set fn_SAVE_PRODUCTS_PACKAGE = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_PRODUCTS_PACKAGE(Optional lngProductID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_PRODUCT_PACKAGE SET " _
    & "ProductID = " & mvarProductID & "," _
    & "PackageID = " & mvarPackageID & "," _
    & "SupplierPrice = " & mvarSupplierPrice & "," _
    & "Qty = " & mvarQty & "," _
    & "SellingPriceWithoutTax = " & mvarSellingPriceWithoutTax & "," _
    & "VAT = " & mvarVAT & "," _
    & "NHIL = " & mvarNHIL & "," _
    & "SellingPriceWithTax = " & mvarSellingPriceWithTax _
    & " WHERE ProductID = " & lngProductID
    
    Set fn_UPDATE_PRODUCTS_PACKAGE = db.fn_MANAGE_RECORD

End Function

Public Function fn_DELETE_PRODUCT_PACKAGE(Optional lngProductID As Long) As ADODB.Recordset

    db.strQueryString = "DELETE FROM tb_PRODUCT_PACKAGE WHERE ProductID = " & lngProductID

    Set fn_DELETE_PRODUCT_PACKAGE = db.fn_MANAGE_RECORD

End Function

