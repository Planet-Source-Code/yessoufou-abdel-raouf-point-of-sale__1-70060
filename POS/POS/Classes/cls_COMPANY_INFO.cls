VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_COMPANY_INFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarCompanyID As Long 'local copy
Private mvarCompanyName As String 'local copy
Private mvarAddress As String 'local copy
Private mvarEMail As String 'local copy
Private mvarPhoneNo As String 'local copy
Private mvarFax As String 'local copy
Private mvarLocation As String 'local copy
Private mvarVATNO As String 'local copy
Private mvarVATRate As Double 'local copy
Private mvarNHILRate As Double 'local copy
Public Property Let NHILRate(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NHILRate = 5
    mvarNHILRate = vData
End Property


Public Property Get NHILRate() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NHILRate
    NHILRate = mvarNHILRate
End Property



Public Property Let VATRate(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VATRate = 5
    mvarVATRate = vData
End Property


Public Property Get VATRate() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VATRate
    VATRate = mvarVATRate
End Property



Public Property Let VATNO(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VATNO = 5
    mvarVATNO = vData
End Property


Public Property Get VATNO() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VATNO
    VATNO = mvarVATNO
End Property



Public Property Let Location(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Location = 5
    mvarLocation = vData
End Property


Public Property Get Location() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Location
    Location = mvarLocation
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



Public Property Let CompanyName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CompanyName = 5
    mvarCompanyName = vData
End Property


Public Property Get CompanyName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CompanyName
    CompanyName = mvarCompanyName
End Property



Public Property Let CompanyID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CompanyID = 5
    mvarCompanyID = vData
End Property


Public Property Get CompanyID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CompanyID
    CompanyID = mvarCompanyID
End Property

Public Function fn_LOAD_COMPANY() As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    db.strQueryString = "SELECT * FROM tb_COMPANY_INFO"


    Set fn_LOAD_COMPANY = db.fn_MANAGE_RECORD

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

Public Function fn_SAVE_COMPANY_INFO() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_COMPANY_INFO " _
    & "(CompanyID,CompanyName,Address,EMail,PhoneNo,Fax,Location,VatNo,VatRate,NHILRate) " _
    & " VALUES " _
    & "(" & fn_AUTOGEN & "," _
    & "'" & mvarCompanyName & "'," _
    & "'" & mvarAddress & "'," _
    & "'" & mvarEMail & "'," _
    & "'" & mvarPhoneNo & "'," _
    & "'" & mvarFax & "'," _
    & "'" & mvarLocation & "'," _
    & "'" & mvarVATNO & "'," _
    & mvarVATRate & "," _
    & mvarNHILRate & ")"

    Set fn_SAVE_COMPANY_INFO = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_COMPANY_INFO(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_COMPANY_INFO SET " _
    & "CompanyName = '" & mvarCompanyName & "'," _
    & "Address = '" & mvarAddress & "'," _
    & "EMail = '" & mvarEMail & "'," _
    & "PhoneNo = '" & mvarPhoneNo & "'," _
    & "Fax = '" & mvarFax & "'," _
    & "Location = '" & mvarLocation & "'," _
    & "VatNo = '" & mvarVATNO & "'," _
    & "VatRate = " & mvarVATRate & "," _
    & "NHILRate = " & mvarNHILRate _
    & " WHERE CompanyID = " & lngID
    
    Set fn_UPDATE_COMPANY_INFO = db.fn_MANAGE_RECORD

End Function


Public Function fn_AUTOGEN() As String

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(CompanyID) FROM tb_COMPANY_INFO"
    
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



