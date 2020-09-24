VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_CATEGORIES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim con_Obj As New ADODB.Connection

'local variable(s) to hold property value(s)
Private mvarCategoryID As Long 'local copy
Private mvarCategoryName As String 'local copy
Private mvarDescription As String 'local copy
Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let CategoryName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CategoryName = 5
    mvarCategoryName = vData
End Property


Public Property Get CategoryName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CategoryName
    CategoryName = mvarCategoryName
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


Public Function fN_LOAD_CATEGORIES(Optional lngCategoryID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset

    If lngCategoryID = 0 Then
        db.strQueryString = "SELECT * FROM tb_CATEGORIES "
        Else
            db.strQueryString = "SELECT * FROM tb_CATEGORIES WHERE CategoryID = " & lngCategoryID
    End If
    
    Set rec = db.fn_MANAGE_RECORD

    Set fN_LOAD_CATEGORIES = rec

End Function

Public Function fN_LOAD_PRODUCTS_REPORT(Optional lngCategoryID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset

    If lngCategoryID = 0 Then
        db.strQueryString = "SELECT * FROM vw_CATEGORIES_PRODUCTS "
        Else
            db.strQueryString = "SELECT * FROM vw_CATEGORIES_PRODUCTS WHERE CategoryID = " & lngCategoryID
    End If
    
    Set rec = db.fn_MANAGE_RECORD

    Set fN_LOAD_PRODUCTS_REPORT = rec

End Function

Public Function fn_LOAD_STOCK_CUMULATIVE_REP() As ADODB.Recordset

    Dim rec As New ADODB.Recordset

    db.strQueryString = "SELECT * FROM vw_STOCK_CUMULATIVE "
    
    Set fn_LOAD_STOCK_CUMULATIVE_REP = db.fn_MANAGE_RECORD


End Function

Public Function fN_LOAD_ACTIVE_PRODUCTS_REPORT(Optional lngCategoryID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset

    If lngCategoryID = 0 Then
        db.strQueryString = "SELECT * FROM vw_CATEGORIES_PRODUCTS WHERE Active = 0"
        Else
            db.strQueryString = "SELECT * FROM vw_CATEGORIES_PRODUCTS WHERE Active = 0 AND CategoryID = " & lngCategoryID
    End If
    
    Set rec = db.fn_MANAGE_RECORD

    Set fN_LOAD_ACTIVE_PRODUCTS_REPORT = rec

End Function

Public Function fN_LOAD_INACTIVE_PRODUCTS_REPORT(Optional lngCategoryID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset

    If lngCategoryID = 0 Then
        db.strQueryString = "SELECT * FROM vw_CATEGORIES_PRODUCTS WHERE Active = 1"
        Else
            db.strQueryString = "SELECT * FROM vw_CATEGORIES_PRODUCTS WHERE Active = 1 AND CategoryID = " & lngCategoryID
    End If
    
    Set rec = db.fn_MANAGE_RECORD

    Set fN_LOAD_INACTIVE_PRODUCTS_REPORT = rec

End Function

Public Function fn_SAVE_CATEGORY_RECORDS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_CATEGORIES " _
    & "(CategoryID,CategoryName,Description) " _
    & " VALUES " _
    & "(" & mvarCategoryID & "," _
    & "'" & mvarCategoryName & "'," _
    & "'" & mvarDescription & "')"

    Set fn_SAVE_CATEGORY_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_CATEGORY_RECORDS(Optional lngCategoryID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_CATEGORIES SET " _
    & "CategoryName = '" & mvarCategoryName & "'," _
    & "Description = '" & mvarDescription & "'" _
    & " WHERE CategoryID = " & lngCategoryID
    
    Set fn_UPDATE_CATEGORY_RECORDS = db.fn_MANAGE_RECORD

End Function

Public Function fn_CHECK_CATEGORY_IN_PRODUCTS(lngID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset

    blnCategoryExist = False

    If lngID = 0 Then
        db.strQueryString = "SELECT * FROM tb_PRODUCTS"
        Else
            db.strQueryString = "SELECT * FROM tb_PRODUCTS WHERE categoryID = " & lngID
    End If

    Set rec = db.fn_MANAGE_RECORD

    If rec.AbsolutePage = -1 Then
        blnCategoryExist = False
        Else
            blnCategoryExist = True
    End If

    Set fn_CHECK_CATEGORY_IN_PRODUCTS = rec

End Function

Public Function fn_DELETE_CATEGORY(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = " DELETE FROM tb_CATEGORIES  " _
    & "Where CategoryID = " & lngID

    Set fn_DELETE_CATEGORY = db.fn_MANAGE_RECORD

End Function

Public Function fn_AUTOGEN() As Long

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(CategoryID) FROM tb_CATEGORIES"
    
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
