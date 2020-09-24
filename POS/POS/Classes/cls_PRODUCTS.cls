VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_PRODUCTS"
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
Private mvarProductName As String 'local copy
Private mvarCategoryID As Long 'local copy
Private mvarActive As Long 'local copy
Private mvarPackageID As Long 'local copy
Private mvarUnitsInStock As Double 'local copy
Private mvarReOrderLevel As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarInitialStock As Long 'local copy
Public Property Let InitialStock(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.InitialStock = 5
    mvarInitialStock = vData
End Property


Public Property Get InitialStock() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.InitialStock
    InitialStock = mvarInitialStock
End Property



Public Property Let ReOrderLevel(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReOrderLevel = 5
    mvarReOrderLevel = vData
End Property


Public Property Get ReOrderLevel() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReOrderLevel
    ReOrderLevel = mvarReOrderLevel
End Property



Public Property Let UnitsInStock(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UnitsInStock = 5
    mvarUnitsInStock = vData
End Property


Public Property Get UnitsInStock() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UnitsInStock
    UnitsInStock = mvarUnitsInStock
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



Public Property Let Active(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Active = 5
    mvarActive = vData
End Property


Public Property Get Active() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Active
    Active = mvarActive
End Property





Public Property Let CategoryID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CategoryID = 5
    mvarCategoryID = vData
End Property


Public Property Get CategoryID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CategoryID
    CategoryID = mvarCategoryID
End Property



Public Property Let ProductName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProductName = 5
    mvarProductName = vData
End Property


Public Property Get ProductName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductName
    ProductName = mvarProductName
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

Public Function fn_LOAD_SUPPLIER_PRODUCT(Optional lngSupplierID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
        
    db.strQueryString = "SELECT tb_SUPPLIER_PRODUCT.ProductID, tb_PRODUCTS.ProductName FROM tb_PRODUCTS INNER JOIN tb_SUPPLIER_PRODUCT ON tb_PRODUCTS.ProductID = tb_SUPPLIER_PRODUCT.ProductID WHERE tb_SUPPLIER_PRODUCT.SupplierID = " & lngSupplierID

    Set rec = db.fn_MANAGE_RECORD
    
    Set fn_LOAD_SUPPLIER_PRODUCT = rec
    
End Function

Public Function fn_LOAD_ACTIVE_PRODUCT(Optional lngSupplierID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
        
    db.strQueryString = "SELECT * FROM vw_SUPPLIER_PRODUCT WHERE Active = 0 AND SupplierID = " & lngSupplierID

    Set rec = db.fn_MANAGE_RECORD
    
    Set fn_LOAD_ACTIVE_PRODUCT = rec
    
End Function

Public Function fn_LOAD_NOT_SUPPLIER_PRODUCT(Optional lngSupplierID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
        
    db.strQueryString = "SELECT tb_SUPPLIER_PRODUCT.ProductID, tb_PRODUCTS.ProductName FROM tb_PRODUCTS INNER JOIN tb_SUPPLIER_PRODUCT ON tb_PRODUCTS.ProductID = tb_SUPPLIER_PRODUCT.ProductID WHERE tb_SUPPLIER_PRODUCT.SupplierID <> " & lngSupplierID

    Set rec = db.fn_MANAGE_RECORD
    
    Set fn_LOAD_NOT_SUPPLIER_PRODUCT = rec
    
End Function

Public Function fN_LOAD_PRODUCTS(Optional lngCategoryID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset

    If lngCategoryID = 0 Then
        db.strQueryString = "SELECT * FROM tb_PRODUCTS "
        Else
            db.strQueryString = "SELECT * FROM tb_PRODUCTS WHERE CategoryID = " & lngCategoryID
    End If
    
    Set rec = db.fn_MANAGE_RECORD

    Set fN_LOAD_PRODUCTS = rec

End Function

Public Function fN_LOAD_PRODUCTS_STOCK(Optional lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset

    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM  vw_CATEGORIES_PRODUCTS WHERE UnitsInStock <= 0 "
        ElseIf lngID = 1 Then
            db.strQueryString = "SELECT * FROM vw_CATEGORIES_PRODUCTS WHERE UnitsInStock = ReorderLevel "
    End If
    
    Set rec = db.fn_MANAGE_RECORD

    Set fN_LOAD_PRODUCTS_STOCK = rec

End Function

Public Function fN_LOAD_ACTIVE_PRODUCTS(Optional lngCategoryID As Long) As ADODB.Recordset

    If lngCategoryID = 0 Then
        db.strQueryString = "SELECT * FROM tb_PRODUCTS WHERE Active = 0"
        Else
            db.strQueryString = "SELECT * FROM tb_PRODUCTS WHERE Active = 0 AND CategoryID = " & lngCategoryID
    End If
    
    Set fN_LOAD_ACTIVE_PRODUCTS = db.fn_MANAGE_RECORD


End Function

Public Function fN_LOAD_PRODUCTS_DETAILS(Optional lngProductID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset

    If lngProductID = 0 Then
        db.strQueryString = "SELECT * FROM tb_PRODUCTS "
        Else
            db.strQueryString = "SELECT * FROM tb_PRODUCTS WHERE ProductID = " & lngProductID
    End If
    
    Set rec = db.fn_MANAGE_RECORD

    Set fN_LOAD_PRODUCTS_DETAILS = rec

End Function

Public Function fn_SAVE_PRODUCTS_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_PRODUCTS " _
    & "(ProductID,ProductName,CategoryID,PackageID,Active,InitialStock,UnitsInStock,ReorderLevel) " _
    & " VALUES " _
    & "(" & mvarProductID & "," _
    & "'" & mvarProductName & "'," _
    & mvarCategoryID & "," _
    & mvarPackageID & "," _
    & mvarActive & "," _
    & mvarInitialStock & "," _
    & mvarUnitsInStock & "," _
    & mvarReOrderLevel & ")"

    Set fn_SAVE_PRODUCTS_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_PRODUCTS_RECORDS(Optional lngProductID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_PRODUCTS SET " _
    & "ProductName = '" & mvarProductName & "'," _
    & "CategoryID = " & mvarCategoryID & "," _
    & "PackageID = " & mvarPackageID & "," _
    & "Active = " & mvarActive & "," _
    & "UnitsInStock = " & mvarUnitsInStock & "," _
    & "InitialStock = " & mvarInitialStock & "," _
    & "ReorderLevel = " & mvarReOrderLevel _
    & " WHERE ProductID = " & lngProductID
    
    Set fn_UPDATE_PRODUCTS_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_SUPPLIER_PRODUCT(Optional lngSupplierID As Long, Optional lngProductID As Long) As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_SUPPLIER_PRODUCT " _
    & "(Supplier_Product_ID,SupplierID,ProductID) " _
    & " VALUES " _
    & "(" & fn_AUTOGEN_SUPPLIER_PRODUCT_ID & "," _
    & lngSupplierID & "," _
    & lngProductID & ")"

    Set fn_UPDATE_SUPPLIER_PRODUCT = db.fn_MANAGE_RECORD

End Function

Public Function fn_CHECK_SUPPLIER_PRODUCT(Optional lngSupplierID As Long, Optional lngProductID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    blnCheckSupplierProductExist = False
    
    db.strQueryString = "SELECT * FROM tb_SUPPLIER_PRODUCT WHERE SupplierID = " & lngSupplierID & " AND ProductID=" & lngProductID

    Set rec = db.fn_MANAGE_RECORD

    If rec.AbsolutePosition <> -1 Then
        blnCheckSupplierProductExist = True
    End If
    
    Set fn_CHECK_SUPPLIER_PRODUCT = rec
    
End Function

Public Function fn_DELETE_SUPPLIER_PRODUCT(Optional lngSupplierID As Long, Optional lngProductID As Long) As ADODB.Recordset

    db.strQueryString = "DELETE FROM tb_SUPPLIER_PRODUCT WHERE SupplierID = " & lngSupplierID & " AND ProductID = " & lngProductID

    Set fn_DELETE_SUPPLIER_PRODUCT = db.fn_MANAGE_RECORD

End Function



Public Function fn_AUTOGEN() As Long

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(ProductID) FROM tb_PRODUCTS"
    
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


Public Function fn_AUTOGEN_SUPPLIER_PRODUCT_ID() As Long

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(Supplier_Product_ID) FROM tb_SUPPLIER_PRODUCT"
    
    Set rec = db.fn_MANAGE_RECORD
    
    If rec.AbsolutePosition <> -1 Then
        If IsNull(rec(0)) Then
            fn_AUTOGEN_SUPPLIER_PRODUCT_ID = 1
        Else
            fn_AUTOGEN_SUPPLIER_PRODUCT_ID = Val(rec(0)) + 1
        End If
    Else
        fn_AUTOGEN_SUPPLIER_PRODUCT_ID = 1
    End If

End Function
