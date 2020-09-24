VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_SALARIES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarSalaryID As Long 'local copy
Private mvarEmployeeID As Long 'local copy
Private mvarBasicPay As Double 'local copy
Private mvarTotalAllowance As Double 'local copy
Private mvarGrossPay As Double 'local copy
Private mvarTotalDeduction As Double 'local copy
Private mvarNetPay As Double 'local copy
Private mvarDA As Double 'local copy
Private mvarHRA As Double 'local copy
Private mvarCCA As Double 'local copy
Private mvarTransport As Double 'local copy
Private mvarGPF As Double 'local copy
Private mvarInssurance As Double 'local copy
Private mvarIncomeTax As Double 'local copy
Private mvarPTax As Double 'local copy

Private mvarSalaryPeriodID As Long 'local copy
Public Property Let SalaryPeriodID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SalaryPeriodID = 5
    mvarSalaryPeriodID = vData
End Property


Public Property Get SalaryPeriodID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SalaryPeriodID
    SalaryPeriodID = mvarSalaryPeriodID
End Property



Public Property Let SalaryID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SalaryID = 5
    mvarSalaryID = vData
End Property


Public Property Get SalaryID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SalaryID
    SalaryID = mvarSalaryID
End Property
Public Property Let PTax(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PTax = 5
    mvarPTax = vData
End Property


Public Property Get PTax() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PTax
    PTax = mvarPTax
End Property



Public Property Let IncomeTax(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IncomeTax = 5
    mvarIncomeTax = vData
End Property


Public Property Get IncomeTax() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IncomeTax
    IncomeTax = mvarIncomeTax
End Property



Public Property Let Inssurance(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Inssurance = 5
    mvarInssurance = vData
End Property


Public Property Get Inssurance() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Inssurance
    Inssurance = mvarInssurance
End Property



Public Property Let GPF(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GPF = 5
    mvarGPF = vData
End Property


Public Property Get GPF() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GPF
    GPF = mvarGPF
End Property



Public Property Let Transport(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Transport = 5
    mvarTransport = vData
End Property


Public Property Get Transport() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Transport
    Transport = mvarTransport
End Property



Public Property Let CCA(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CCA = 5
    mvarCCA = vData
End Property


Public Property Get CCA() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CCA
    CCA = mvarCCA
End Property



Public Property Let HRA(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HRA = 5
    mvarHRA = vData
End Property


Public Property Get HRA() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HRA
    HRA = mvarHRA
End Property



Public Property Let DA(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DA = 5
    mvarDA = vData
End Property


Public Property Get DA() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DA
    DA = mvarDA
End Property



Public Property Let NetPay(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NetPay = 5
    mvarNetPay = vData
End Property


Public Property Get NetPay() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NetPay
    NetPay = mvarNetPay
End Property



Public Property Let TotalDeduction(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalDeduction = 5
    mvarTotalDeduction = vData
End Property


Public Property Get TotalDeduction() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalDeduction
    TotalDeduction = mvarTotalDeduction
End Property



Public Property Let GrossPay(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GrosPay = 5
    mvarGrossPay = vData
End Property


Public Property Get GrossPay() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GrosPay
    GrossPay = mvarGrossPay
End Property



Public Property Let TotalAllowance(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalAllowance = 5
    mvarTotalAllowance = vData
End Property


Public Property Get TotalAllowance() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalAllowance
    TotalAllowance = mvarTotalAllowance
End Property



Public Property Let BasicPay(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BasicPay = 5
    mvarBasicPay = vData
End Property


Public Property Get BasicPay() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BasicPay
    BasicPay = mvarBasicPay
End Property



Public Property Let EmployeeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EmployeeID = 5
    mvarEmployeeID = vData
End Property


Public Property Get EmployeeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EmployeeID
    EmployeeID = mvarEmployeeID
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

Public Function fn_LOAD_SALARIES(Optional lngEmployeeID As Long) As ADODB.Recordset
        
'    db.strQueryString = "SELECT tb_EMPLOYEES.EmployeeID, tb_SALARIES.BasicPay, tb_SALARIES.TotalAllowance, tb_SALARIES.GrossPay,tb_SALARIES.TotalDeduction,tb_SALARIES.NetPay FROM tb_EMPLOYEES INNER JOIN tb_SALARIES ON tb_EMPLOYEES.EmployeeID = tb_SALARIES.EmployeeID WHERE tb_SALARIES.EmployeeID = " & lngEmployeeID
    db.strQueryString = "SELECT * FROM vw_SALARIES WHERE employeeID=" & lngEmployeeID
    
    Set fn_LOAD_SALARIES = db.fn_MANAGE_RECORD
    
End Function


Public Function fn_SAVE_SALARIES(Optional lngSalaryID As Long, Optional lngEmployeeID As Long) As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_SALARIES " _
    & "(SalaryID,EmployeeID,BasicPay,TotalAllowance,GrossPay,TotalDeduction,NetPay) " _
    & " VALUES " _
    & "(" & lngSalaryID & "," _
    & lngEmployeeID & "," _
    & mvarBasicPay & "," _
    & mvarTotalAllowance & "," _
    & mvarGrossPay & "," _
    & mvarTotalDeduction & "," _
    & mvarNetPay & ")"

    Set fn_SAVE_SALARIES = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_SALARIES(Optional lngEmployeeID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_SALARIES SET " _
    & "BasicPay = " & mvarBasicPay & "," _
    & "TotalAllowance = " & mvarTotalAllowance & "," _
    & "GrossPay = " & mvarGrossPay & "," _
    & "TotalDeduction = " & mvarTotalDeduction & "," _
    & "NetPay = " & mvarNetPay _
    & " WHERE EmployeeID = " & lngEmployeeID
    
    Set fn_UPDATE_SALARIES = db.fn_MANAGE_RECORD

End Function

Public Function fn_SAVE_ALLOWANCE_DETAILS(Optional lngSalaryID As Long) As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_ALLOWANCE_DETAILS " _
    & "(SalaryID,DA,HRA,CCA,Transport) " _
    & " VALUES " _
    & "(" & lngSalaryID & "," _
    & mvarDA & "," _
    & mvarHRA & "," _
    & mvarCCA & "," _
    & mvarTransport & ")"

    Set fn_SAVE_ALLOWANCE_DETAILS = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_ALLOWANCE_DETAILS(Optional lngSalaryID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_ALLOWANCE_DETAILS SET " _
    & "SalaryID = " & mvarSalaryID & "," _
    & "DA = " & mvarDA & "," _
    & "HRA = " & mvarHRA & "," _
    & "CCA = " & mvarCCA & "," _
    & "Transport = " & mvarTransport _
    & " WHERE SalaryID = " & lngSalaryID
    
    Set fn_UPDATE_ALLOWANCE_DETAILS = db.fn_MANAGE_RECORD

End Function
Public Function fn_SAVE_DEDUCTION_DETAILS(Optional lngSalaryID As Long) As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_DEDUCTION_DETAILS " _
    & "(SalaryID,GPF,Inssurance,IncomeTax,PTax) " _
    & " VALUES " _
    & "(" & lngSalaryID & "," _
    & mvarGPF & "," _
    & mvarInssurance & "," _
    & mvarIncomeTax & "," _
    & mvarPTax & ")"

    Set fn_SAVE_DEDUCTION_DETAILS = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_DEDUCTION_DETAILS(Optional lngSalaryID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_DEDUCTION_DETAILS SET " _
    & "SalaryID = '" & mvarSalaryID & "'," _
    & "GPF = " & mvarGPF & "," _
    & "Inssurance = '" & mvarInssurance & "'," _
    & "IncomeTax = " & mvarIncomeTax & "," _
    & "PTax = " & mvarPTax _
    & " WHERE SalaryID = " & lngSalaryID
    
    Set fn_UPDATE_DEDUCTION_DETAILS = db.fn_MANAGE_RECORD

End Function

Public Function fn_CHECK_IF_PAID(lngSalaryID As Long, lngSalaryPeriodID As Long)
    
    blnCheckIfSalaryPaid = True
    
    db.strQueryString = "SELECT * FROM tb_SALARY_PERIOD_DETAILS WHERE SalaryID = " & lngSalaryID & " AND SalaryPeriodID = " & lngSalaryPeriodID
    
    Set fn_CHECK_IF_PAID = db.fn_MANAGE_RECORD
    
    If fn_CHECK_IF_PAID.AbsolutePosition <> -1 Then
        blnCheckIfSalaryPaid = False
    End If

End Function

Public Function fn_SAVE_SALARY_PERIOD_DETAILS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_SALARY_PERIOD_DETAILS " _
    & "(SalaryPeriodID,SalaryID) " _
    & " VALUES " _
    & "(" & mvarSalaryPeriodID & "," _
    & mvarSalaryID & ")"

    Set fn_SAVE_SALARY_PERIOD_DETAILS = db.fn_MANAGE_RECORD

End Function

Public Function fn_AUTOGEN() As Long
    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(SalaryID) FROM tb_SALARIES"
    
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

