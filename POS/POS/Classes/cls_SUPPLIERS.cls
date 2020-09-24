VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_SUPPLIERS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarSupplierID As Integer 'local copy
Private mvarCompanyName As String 'local copy
Private mvarContactName As Variant 'local copy
Private mvarContactTitle As Integer 'local copy
Private mvarAddress As String 'local copy
Private mvarCity As String 'local copy
Private mvarRegion As String 'local copy
Private mvarPostalCode As String 'local copy
Private mvarCountry As String 'local copy
Private mvarPhone As String 'local copy
Private mvarFax As String 'local copy
Private mvarHomePage As String 'local copy
Public Property Let HomePage(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HomePage = 5
    mvarHomePage = vData
End Property


Public Property Get HomePage() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HomePage
    HomePage = mvarHomePage
End Property



Public Property Let Fax(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Fax = 5
    mvarFax = vData
End Property


Public Property Get Fax() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Fax
    Fax = mvarFax
End Property



Public Property Let Phone(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Phone = 5
    mvarPhone = vData
End Property


Public Property Get Phone() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Phone
    Phone = mvarPhone
End Property



Public Property Let Country(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Country = 5
    mvarCountry = vData
End Property


Public Property Get Country() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Country
    Country = mvarCountry
End Property



Public Property Let PostalCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PostalCode = 5
    mvarPostalCode = vData
End Property


Public Property Get PostalCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PostalCode
    PostalCode = mvarPostalCode
End Property



Public Property Let Region(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Region = 5
    mvarRegion = vData
End Property


Public Property Get Region() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Region
    Region = mvarRegion
End Property



Public Property Let City(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.City = 5
    mvarCity = vData
End Property


Public Property Get City() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.City
    City = mvarCity
End Property



Public Property Let Address(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Address = 5
    mvarAddress = vData
End Property


Public Property Get Address() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Address
    Address = mvarAddress
End Property



Public Property Let ContactTitle(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ContactTitle = 5
    mvarContactTitle = vData
End Property


Public Property Get ContactTitle() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ContactTitle
    ContactTitle = mvarContactTitle
End Property



Public Property Let ContactName(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ContactName = 5
    mvarContactName = vData
End Property


Public Property Set ContactName(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ContactName = Form1
    Set mvarContactName = vData
End Property


Public Property Get ContactName() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ContactName
    If IsObject(mvarContactName) Then
        Set ContactName = mvarContactName
    Else
        ContactName = mvarContactName
    End If
End Property



Public Property Let CompanyName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SupplierName = 5
    mvarCompanyName = vData
End Property


Public Property Get CompanyName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SupplierName
    CompanyName = mvarCompanyName
End Property



Public Property Let SupplierID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SupplierID = 5
    mvarSupplierID = vData
End Property


Public Property Get SupplierID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SupplierID
    SupplierID = mvarSupplierID
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

Public Function fn_LOAD_SUPPLIERS(Optional lngSupplierID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset

    If lngSupplierID = 0 Then
        db.strQueryString = "SELECT * FROM tb_SUPPLIERS "
        Else
            db.strQueryString = "SELECT * FROM tb_SUPPLIERS WHERE SupplierID = " & lngSupplierID
    End If
    
    Set rec = db.fn_MANAGE_RECORD
    

    Set fn_LOAD_SUPPLIERS = rec

End Function


Public Function fn_LOAD_SUPPLIERS_REPORT(Optional lngSupplierID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset

    If lngSupplierID = 0 Then
        db.strQueryString = "SELECT * FROM vw_SUPPLIERS_AND_PRODUCTS "
        Else
            db.strQueryString = "SELECT * FROM vw_SUPPLIERS_AND_PRODUCTS WHERE SupplierID = " & lngSupplierID
    End If
    
    Set rec = db.fn_MANAGE_RECORD
    

    Set fn_LOAD_SUPPLIERS_REPORT = rec

End Function


Public Function fn_SAVE_SUPPLIERS_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_SUPPLIERS " _
    & "(SupplierID,CompanyName,ContactName,ContactTitle,Address,City,Region,PostalCode,Country,Phone,Fax,HomePage) " _
    & " VALUES " _
    & "(" & mvarSupplierID & "," _
    & "'" & mvarCompanyName & "'," _
    & "'" & mvarContactName & "'," _
    & mvarContactTitle & "," _
    & "'" & mvarAddress & "'," _
    & "'" & mvarCity & "'," _
    & "'" & mvarRegion & "'," _
    & "'" & mvarPostalCode & "'," _
    & "'" & mvarCountry & "'," _
    & "'" & mvarPhone & "'," _
    & "'" & mvarFax & "'," _
    & "'" & mvarHomePage & "')"

    Set fn_SAVE_SUPPLIERS_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_SUPPLIERS_RECORDS(Optional lngSupplierID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_SUPPLIERS SET " _
    & "CompanyName = '" & mvarCompanyName & "'," _
    & "ContactName = '" & mvarContactName & "'," _
    & "ContactTitle = " & mvarContactTitle & "," _
    & "Address = '" & mvarAddress & "'," _
    & "City = '" & mvarCity & "'," _
    & "Region = '" & mvarRegion & "'," _
    & "PostalCode = '" & mvarPostalCode & "'," _
    & "Country = '" & mvarCountry & "'," _
    & "Phone = '" & mvarPhone & "'," _
    & "Fax = '" & mvarFax & "'," _
    & "HomePage = '" & mvarHomePage & "'" _
    & " WHERE SupplierID = " & lngSupplierID
    
    Set fn_UPDATE_SUPPLIERS_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_CHECK_SUPPLIER_IN_ORDERS(lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset

    blnSupplierExist = False

    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM tb_ORDERS"
        Else
            db.strQueryString = "SELECT * FROM tb_ORDERS WHERE SupplierID = " & lngID
    End If

    Set rec = db.fn_MANAGE_RECORD

    If rec.AbsolutePage = -1 Then
        blnSupplierExist = False
        Else
            blnSupplierExist = True
    End If

    Set fn_CHECK_SUPPLIER_IN_ORDERS = rec

End Function

Public Function fn_DELETE_SUPPLIER(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = " DELETE FROM tb_SUPPLIERS  " _
    & "Where SupplierID = " & lngID

    Set fn_DELETE_SUPPLIER = db.fn_MANAGE_RECORD

End Function

Public Function fn_AUTOGEN() As Long

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(SupplierID) FROM tb_SUPPLIERS"
    
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

