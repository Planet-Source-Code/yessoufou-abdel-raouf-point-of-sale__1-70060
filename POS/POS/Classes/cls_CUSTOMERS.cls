VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_CUSTOMERS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarCustomerID As Long 'local copy
Private mvarCustomerNo As String 'local copy
Private mvarFirstName As String 'local copy
Private mvarOtherNames As String 'local copy
Private mvarTitle As Long 'local copy
Private mvarGender As String 'local copy
Private mvarPhoneNo As String 'local copy
Private mvarEMail As String 'local copy
Private mvarAddress As String 'local copy
Private mvarPicture As String 'local copy
'local variable(s) to hold property value(s)
Private mvarLastName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarDateOfBirth As Date 'local copy
'local variable(s) to hold property value(s)
Private mvarFinalDestination As Variant 'local copy
Public Property Let FinalDestination(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FinalDestination = 5
    mvarFinalDestination = vData
End Property


Public Property Set FinalDestination(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.FinalDestination = Form1
    Set mvarFinalDestination = vData
End Property


Public Property Get FinalDestination() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FinalDestination
    If IsObject(mvarFinalDestination) Then
        Set FinalDestination = mvarFinalDestination
    Else
        FinalDestination = mvarFinalDestination
    End If
End Property



Public Property Let DateOfBirth(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateOfBirth = 5
    mvarDateOfBirth = vData
End Property


Public Property Get DateOfBirth() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateOfBirth
    DateOfBirth = mvarDateOfBirth
End Property



Public Property Let LastName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastName = 5
    mvarLastName = vData
End Property


Public Property Get LastName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastName
    LastName = mvarLastName
End Property



Public Property Let Picture(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Picture = 5
    mvarPicture = vData
End Property


Public Property Get Picture() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Picture
    Picture = mvarPicture
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



Public Property Let EMail(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EMail = 5
    mvarEMail = vData
End Property


Public Property Get EMail() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EMail
    EMail = mvarEMail
End Property



Public Property Let PhoneNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PhoneNo = 5
    mvarPhoneNo = vData
End Property


Public Property Get PhoneNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PhoneNo
    PhoneNo = mvarPhoneNo
End Property



Public Property Let Gender(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Gender = 5
    mvarGender = vData
End Property


Public Property Get Gender() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Gender
    Gender = mvarGender
End Property



Public Property Let Title(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Title = 5
    mvarTitle = vData
End Property


Public Property Get Title() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Title
    Title = mvarTitle
End Property



Public Property Let OtherNames(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OtherNames = 5
    mvarOtherNames = vData
End Property


Public Property Get OtherNames() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OtherNames
    OtherNames = mvarOtherNames
End Property



Public Property Let FirstName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FirstName = 5
    mvarFirstName = vData
End Property


Public Property Get FirstName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FirstName
    FirstName = mvarFirstName
End Property



Public Property Let CustomerNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CustomerNo = 5
    mvarCustomerNo = vData
End Property


Public Property Get CustomerNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CustomerNo
    CustomerNo = mvarCustomerNo
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

Public Function fn_LOAD_CUSTOMERS(lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM tb_CUSTOMERS"
        Else
            db.strQueryString = "SELECT * FROM tb_CUSTOMERS WHERE customerID = " & lngID
    End If

    Set rec = db.fn_MANAGE_RECORD

    Set fn_LOAD_CUSTOMERS = rec

End Function

Public Function fn_LOAD_CUSTOMERS_REPORT(lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM vw_CUSTOMERS"
        Else
            db.strQueryString = "SELECT * FROM vw_CUSTOMERS WHERE customerID = " & lngID
    End If

    Set rec = db.fn_MANAGE_RECORD

    Set fn_LOAD_CUSTOMERS_REPORT = rec

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

Public Function fn_SAVE_CUSTOMER_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_CUSTOMERS " _
    & "(CustomerID,CustomerNo,FirstName,LastName,OtherNames,Title,Gender,PhoneNo,EMail,Address,Picture) " _
    & " VALUES " _
    & "(" & mvarCustomerID & "," _
    & "'" & mvarCustomerNo & "'," _
    & "'" & mvarFirstName & "'," _
    & "'" & mvarLastName & "'," _
    & "'" & mvarOtherNames & "'," _
    & mvarTitle & "," _
    & "'" & mvarGender & "'," _
    & "'" & mvarPhoneNo & "'," _
    & "'" & mvarEMail & "'," _
    & "'" & mvarAddress & "'," _
    & "'" & mvarPicture & "')"

    Set fn_SAVE_CUSTOMER_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_CUSTOMER_RECORDS(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_CUSTOMERS SET " _
    & "CustomerNo = '" & mvarCustomerNo & "'," _
    & "FirstName = '" & mvarFirstName & "'," _
    & "LastName = '" & mvarLastName & "'," _
    & "OtherNames = '" & mvarOtherNames & "'," _
    & "Title = " & mvarTitle & "," _
    & "Gender = '" & mvarGender & "'," _
    & "PhoneNo = '" & mvarPhoneNo & "'," _
    & "EMail = '" & mvarEMail & "'," _
    & "Address = '" & mvarAddress & "'," _
    & "Picture = '" & Picture & "'" _
    & " WHERE CustomerID = " & lngID
    
    Set fn_UPDATE_CUSTOMER_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_CHECK_CUSTOMER_IN_SALES(lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset

    blnCustomerExist = False

    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM tb_SALES"
        Else
            db.strQueryString = "SELECT * FROM tb_SALES WHERE customerID = " & lngID
    End If

    Set rec = db.fn_MANAGE_RECORD

    If rec.AbsolutePage = -1 Then
        blnCustomerExist = False
        Else
            blnCustomerExist = True
    End If

    Set fn_CHECK_CUSTOMER_IN_SALES = rec

End Function

Public Function fn_DELETE_CUSTOMER(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = " DELETE FROM tb_CUSTOMERS  " _
    & "Where CustomerID = " & lngID

    Set fn_DELETE_CUSTOMER = db.fn_MANAGE_RECORD

End Function

Public Function fn_DELETE_PICTURE(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = " UPDATE tb_CUSTOMERS SET " _
    & "Picture =  '" & "" & "'" _
    & "Where CustomerID = " & lngID

    Set fn_DELETE_PICTURE = db.fn_MANAGE_RECORD

End Function



Public Function fn_AUTOGEN() As String

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(CustomerNo) FROM tb_CUSTOMERS"
    
    Set rec = db.fn_MANAGE_RECORD
    
    If rec.AbsolutePosition <> -1 Then
        If IsNull(rec(0)) Then
            fn_AUTOGEN = "C00001"
        Else
            fn_AUTOGEN = "C" & Format(Right(Trim(rec(0)), 5) + 1, "00000")
        End If
    Else
        fn_AUTOGEN = "C00001"
    End If

End Function

Public Function fn_ID_AUTOGEN() As Long

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(CustomerID) FROM tb_CUSTOMERS"
    
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
