VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_EMPLOYEES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim con_Obj As New ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarEmployeeID As Long 'local copy
Private mvarLastName As String 'local copy
Private mvarFirstName As String 'local copy
Private mvarTitle As String 'local copy
Private mvarTitleOfCourtesy As String 'local copy
Private mvarBirthDate As Date 'local copy
Private mvarHireDate As Date 'local copy
Private mvarAddress As String 'local copy
Private mvarCity As String 'local copy
Private mvarRegion As String 'local copy
Private mvarPostalCode As String 'local copy
Private mvarCountry As String 'local copy
Private mvarHomePhone As String 'local copy
Private mvarExtension As String 'local copy
Private mvarPhoto As String 'local copy
Private mvarNotes As String 'local copy
Private mvarReportsTo As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarWorkingStatus As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarEmployeeNo As String 'local copy
Public Property Let EmployeeNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EmployeeNo = 5
    mvarEmployeeNo = vData
End Property


Public Property Get EmployeeNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EmployeeNo
    EmployeeNo = mvarEmployeeNo
End Property



Public Property Let WorkingStatus(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WorkingStatus = 5
    mvarWorkingStatus = vData
End Property


Public Property Get WorkingStatus() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WorkingStatus
    WorkingStatus = mvarWorkingStatus
End Property



Public Property Let ReportsTo(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReportsTo = 5
    mvarReportsTo = vData
End Property


Public Property Get ReportsTo() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReportsTo
    ReportsTo = mvarReportsTo
End Property



Public Property Let Notes(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Notes = 5
    mvarNotes = vData
End Property


Public Property Get Notes() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Notes
    Notes = mvarNotes
End Property



Public Property Let Photo(ByVal vData As String)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax:  x.Photo = Form1
    mvarPhoto = vData
End Property


Public Property Get Photo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Photo
    Photo = mvarPhoto
End Property



Public Property Let Extension(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Extension = 5
    mvarExtension = vData
End Property


Public Property Get Extension() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Extension
    Extension = mvarExtension
End Property



Public Property Let HomePhone(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HomePhone = 5
    mvarHomePhone = vData
End Property


Public Property Get HomePhone() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HomePhone
    HomePhone = mvarHomePhone
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



Public Property Let HireDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HireDate = 5
    mvarHireDate = vData
End Property


Public Property Get HireDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HireDate
    HireDate = mvarHireDate
End Property



Public Property Let BirthDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BirthDate = 5
    mvarBirthDate = vData
End Property


Public Property Get BirthDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BirthDate
    BirthDate = mvarBirthDate
End Property



Public Property Let TitleOfCourtesy(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TitleOfCourtesy = 5
    mvarTitleOfCourtesy = vData
End Property


Public Property Get TitleOfCourtesy() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TitleOfCourtesy
    TitleOfCourtesy = mvarTitleOfCourtesy
End Property



Public Property Let Title(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Title = 5
    mvarTitle = vData
End Property


Public Property Get Title() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Title
    Title = mvarTitle
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
Public Function fn_LOAD_EMPLOYEES(Optional lngEmployeeID As Long) As ADODB.Recordset

    If lngEmployeeID = 0 Then
        db.strQueryString = "SELECT * FROM tb_EMPLOYEES "
        Else
            db.strQueryString = "SELECT * FROM tb_EMPLOYEES WHERE EmployeeID = " & lngEmployeeID
    End If

    Set fn_LOAD_EMPLOYEES = db.fn_MANAGE_RECORD

End Function

Public Function fn_SAVE_EMPLOYEE_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_EMPLOYEES " _
    & "(EmployeeNo,FirstName,LastName,Title,TitleOfCourtesy,BirthDate,HireDate,Address,City,Region,PostalCode,Country,HomePhone,Extension,Notes,ReportsTo,WorkingStatus,Photo) " _
    & " VALUES " _
    & "('" & mvarEmployeeNo & "'," _
    & "'" & mvarFirstName & "'," _
    & "'" & mvarLastName & "'," _
    & "'" & mvarTitle & "'," _
    & mvarTitleOfCourtesy & "," _
    & "'" & mvarBirthDate & "'," _
    & "'" & mvarHireDate & "'," _
    & "'" & mvarAddress & "'," _
    & "'" & mvarCity & "'," _
    & "'" & mvarRegion & "'," _
    & "'" & mvarPostalCode & "'," _
    & "'" & mvarCountry & "'," _
    & "'" & mvarHomePhone & "'," _
    & "'" & mvarExtension & "'," _
    & "'" & mvarNotes & "'," _
    & mvarReportsTo & "," _
    & mvarWorkingStatus & "," _
    & "'" & mvarPhoto & "')"

    Set fn_SAVE_EMPLOYEE_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_EMPLOYEE_RECORDS(Optional lngEmployeeID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_EMPLOYEES SET " _
    & "FirstName = '" & mvarFirstName & "'," _
    & "LastName = '" & mvarLastName & "'," _
    & "Title = '" & mvarTitle & "'," _
    & "TitleOfCourtesy = " & mvarTitleOfCourtesy & "," _
    & "BirthDate = '" & mvarBirthDate & "'," _
    & "HireDate = '" & mvarHireDate & "'," _
    & "Address = '" & mvarAddress & "'," _
    & "City = '" & mvarCity & "'," _
    & "Region = '" & mvarRegion & "'," _
    & "PostalCode = '" & mvarPostalCode & "'," _
    & "Country = '" & mvarCountry & "'," _
    & "HomePhone = '" & mvarHomePhone & "'," _
    & "Extension = '" & mvarExtension & "'," _
    & "Notes = '" & mvarNotes & "'," _
    & "ReportsTo = " & mvarReportsTo & "," _
    & "WorkingStatus = " & mvarWorkingStatus & "," _
    & "Photo = '" & mvarPhoto & "'" _
    & " WHERE EmployeeID = " & lngEmployeeID
    
    Set fn_UPDATE_EMPLOYEE_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_DELETE_EMPLOYEE_RECORDS(Optional lngEmployeeID As Long) As ADODB.Recordset

    db.strQueryString = " UPDATE tb_EMPLOYEES SET " _
    & "WorkingStatus = 1 " _
    & "Where EmployeeID = " & lngEmployeeID

    Set fn_DELETE_EMPLOYEE_RECORDS = db.fn_MANAGE_RECORD

End Function


Public Function fn_DELETE_EMPLOYEE_PICTURE(Optional lngEmployeeID As Long) As ADODB.Recordset

    db.strQueryString = " UPDATE tb_EMPLOYEES SET " _
    & "Photo =  '" & "" & "'" _
    & "Where EmployeeID = " & lngEmployeeID

    Set fn_DELETE_EMPLOYEE_PICTURE = db.fn_MANAGE_RECORD

End Function

Public Function fn_AUTOGEN() As String
 
    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(EmployeeNo) FROM tb_EMPLOYEES"
    
    Set rec = db.fn_MANAGE_RECORD
    
    If rec.AbsolutePosition <> -1 Then
        If IsNull(rec(0)) Then
            fn_AUTOGEN = "E00001"
        Else
            fn_AUTOGEN = "E" & Format(Right(Trim(rec(0)), 5) + 1, "00000")
        End If
    Else
        fn_AUTOGEN = "E00001"
    End If

End Function
