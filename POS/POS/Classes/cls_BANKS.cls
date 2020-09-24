VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_BANKS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim con_Obj As New ADODB.Connection

'local variable(s) to hold property value(s)
Private mvarBankID As Long 'local copy
Private mvarShortName As String 'local copy
Private mvarBankName As String 'local copy
Private mvarBranch As String 'local copy
Private mvarAccountID As Long 'local copy
Private mvarAccountNo As String 'local copy
Private mvarOpenDate As Date 'local copy
Private mvarAccountTypeID As Long 'local copy
Private mvarAccountName As String 'local copy
Private mvarOpenBalance As Double 'local copy
Private mvarAccountType As String 'local copy
Public Property Let AccountType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccountType = 5
    mvarAccountType = vData
End Property


Public Property Get AccountType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccountType
    AccountType = mvarAccountType
End Property



Public Property Let OpenBalance(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OpenBalance = 5
    mvarOpenBalance = vData
End Property


Public Property Get OpenBalance() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OpenBalance
    OpenBalance = mvarOpenBalance
End Property



Public Property Let AccountName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccountName = 5
    mvarAccountName = vData
End Property


Public Property Get AccountName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccountName
    AccountName = mvarAccountName
End Property



Public Property Let AccountTypeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccountTypeID = 5
    mvarAccountTypeID = vData
End Property


Public Property Get AccountTypeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccountTypeID
    AccountTypeID = mvarAccountTypeID
End Property



Public Property Let OpenDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OpenDate = 5
    mvarOpenDate = vData
End Property


Public Property Get OpenDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OpenDate
    OpenDate = mvarOpenDate
End Property



Public Property Let AccountNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccountNo = 5
    mvarAccountNo = vData
End Property


Public Property Get AccountNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccountNo
    AccountNo = mvarAccountNo
End Property



Public Property Let AccountID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AccountID = 5
    mvarAccountID = vData
End Property


Public Property Get AccountID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AccountID
    AccountID = mvarAccountID
End Property



Public Property Let Branch(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Branch = 5
    mvarBranch = vData
End Property


Public Property Get Branch() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Branch
    Branch = mvarBranch
End Property



Public Property Let BankName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BankName = 5
    mvarBankName = vData
End Property


Public Property Get BankName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BankName
    BankName = mvarBankName
End Property



Public Property Let ShortName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ShortName = 5
    mvarShortName = vData
End Property


Public Property Get ShortName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ShortName
    ShortName = mvarShortName
End Property



Public Property Let BankID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BankID = 5
    mvarBankID = vData
End Property


Public Property Get BankID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BankID
    BankID = mvarBankID
End Property

Public Function fn_LOAD_BANKS(lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    If lngID = 0 Then
        cls_DATABASE_Obj.strQueryString = "SELECT * FROM tb_BANK"
        Else
            cls_DATABASE_Obj.strQueryString = "SELECT * FROM tb_BANK WHERE BankID = " & lngID
    End If

    Set rec = cls_DATABASE_Obj.fn_MANAGE_RECORD

    Set fn_LOAD_BANKS = rec

End Function

Public Function fn_LOAD_BANKS_ACCOUNTS(lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    If lngID = 0 Then
        cls_DATABASE_Obj.strQueryString = "SELECT * FROM tb_BANK_ACCOUNT"
        Else
            cls_DATABASE_Obj.strQueryString = "SELECT * FROM tb_BANK_ACCOUNT WHERE BankID = " & lngID
    End If

    Set rec = cls_DATABASE_Obj.fn_MANAGE_RECORD

    Set fn_LOAD_BANKS_ACCOUNTS = rec

End Function

Public Function fn_LOAD_BANKS_ACCOUNTS_RECORDS(lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    If lngID = 0 Then
        cls_DATABASE_Obj.strQueryString = "SELECT * FROM tb_BANK_ACCOUNT"
        Else
            cls_DATABASE_Obj.strQueryString = "SELECT * FROM tb_BANK_ACCOUNT WHERE AccountID = " & lngID
    End If

    Set rec = cls_DATABASE_Obj.fn_MANAGE_RECORD

    Set fn_LOAD_BANKS_ACCOUNTS_RECORDS = rec

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

Public Function fn_SAVE_BANK_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_BANK " _
    & "(ShortName,BankName,Branch) " _
    & " VALUES " _
    & "('" & mvarShortName & "'," _
    & "'" & mvarBankName & "'," _
    & "'" & mvarBranch & "')"

    Set fn_SAVE_BANK_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_SAVE_BANK_ACCOUNT_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_BANK_ACCOUNT " _
    & "(AccountNo,OpenDate,BankID,AccountTypeID,AccountName,OpenBalance) " _
    & " VALUES " _
    & "('" & mvarAccountNo & "'," _
    & "'" & mvarOpenDate & "'," _
    & mvarBankID & "," _
    & mvarAccountTypeID & "," _
    & "'" & mvarAccountName & "'," _
    & mvarOpenBalance & ")"

    Set fn_SAVE_BANK_ACCOUNT_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_BANK_RECORDS(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_BANK SET " _
    & "ShortName = '" & mvarShortName & "'," _
    & "BankName = '" & mvarBankName & "'," _
    & "Branch = '" & mvarBranch & "'" _
    & " WHERE BankID = " & lngID
    
    Set fn_UPDATE_BANK_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_BANK_ACCOUNT_RECORDS(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_BANK_ACCOUNT SET " _
    & "AccountNo = '" & mvarAccountNo & "'," _
    & "OpenDate = '" & mvarOpenDate & "'," _
    & "BankID = " & mvarBankID & "," _
    & "AccountTypeID = " & mvarAccountTypeID & "," _
    & "AccountName = '" & mvarAccountName & "'," _
    & "OpenBalance = " & mvarOpenBalance & "" _
    & " WHERE AccountID = " & lngID
    
    Set fn_UPDATE_BANK_ACCOUNT_RECORDS = db.fn_MANAGE_RECORD

End Function

