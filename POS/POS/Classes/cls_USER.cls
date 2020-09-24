VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_USER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Dim con_Obj As ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarUserID As Long 'local copy
Private mvarUserName As String 'local copy
Private mvarPassword As String 'local copy
Private mvarFullName As String 'local copy
Private mvarStatus As Long 'local copy
Private mvarDeleted As Long 'local copy
Private mvarActive As Long 'local copy
Private mvarAdmin As Long 'local copy
Private mvarLastPasswordChangeDate As Date 'local copy
Private mvarLastLoginDate As Date 'local copy
Private mvarLastLoginWorkStation As String 'local copy
'local variable(s) to hold property value(s)
Private mvarUsersFeaturesID As Long 'local copy
Public Property Let UsersFeaturesID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UsersFeaturesID = 5
    mvarUsersFeaturesID = vData
End Property


Public Property Get UsersFeaturesID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UsersFeaturesID
    UsersFeaturesID = mvarUsersFeaturesID
End Property






Public Property Let LastLoginWorkStation(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastLoginWorkStation = 5
    mvarLastLoginWorkStation = vData
End Property


Public Property Get LastLoginWorkStation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastLoginWorkStation
    LastLoginWorkStation = mvarLastLoginWorkStation
End Property



Public Property Let LastLoginDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastLoginDate = 5
    mvarLastLoginDate = vData
End Property


Public Property Get LastLoginDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastLoginDate
    LastLoginDate = mvarLastLoginDate
End Property



Public Property Let LastPasswordChangeDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastPasswordChangeDate = 5
    mvarLastPasswordChangeDate = vData
End Property


Public Property Get LastPasswordChangeDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastPasswordChangeDate
    LastPasswordChangeDate = mvarLastPasswordChangeDate
End Property



Public Property Let Admin(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Admin = 5
    mvarAdmin = vData
End Property


Public Property Get Admin() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Admin
    Admin = mvarAdmin
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



Public Property Let Deleted(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Deleted = 5
    mvarDeleted = vData
End Property


Public Property Get Deleted() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Deleted
    Deleted = mvarDeleted
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



Public Property Let FullName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FullName = 5
    mvarFullName = vData
End Property


Public Property Get FullName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FullName
    FullName = mvarFullName
End Property



Public Property Let Password(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Password = 5
    mvarPassword = vData
End Property


Public Property Get Password() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Password
    Password = mvarPassword
End Property



Public Property Let UserName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserName = 5
    mvarUserName = vData
End Property


Public Property Get UserName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserName
    UserName = mvarUserName
End Property



Public Property Let UserID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserID = 5
    mvarUserID = vData
End Property


Public Property Get UserID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserID
    UserID = mvarUserID
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
    Call MdlFunctions.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub

Public Function fn_CHECK_USER_LOGIN(strUserName As String, strPassword As String) As Boolean

    Dim rec As New ADODB.Recordset
    cls_DATABASE_Obj.strQueryString = "SELECT * FROM tb_USERS WHERE UserName = '" & strUserName & "' AND Password = '" & strPassword & "'"

    Set rec = cls_DATABASE_Obj.fn_MANAGE_RECORD
    
    fn_CHECK_USER_LOGIN = False
    
    If rec.AbsolutePosition <> -1 Then
        fn_CHECK_USER_LOGIN = True
        mvarUserID = rec!UserID
        mvarUserName = rec!UserName
        mvarPassword = rec!Password
        mvarFullName = Trim(rec!FullName) & ""
        mvarAdmin = Trim(rec!Admin)
    End If

End Function

Public Function fn_LOAD_USERS(lngUserID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    If lngUserID = 0 Then
        db.strQueryString = "SELECT * FROM tb_USERS WHERE Deleted = 0"
        Else
            db.strQueryString = "SELECT * FROM tb_USERS WHERE UserID = " & lngUserID
    End If
    Set rec = db.fn_MANAGE_RECORD
    Set fn_LOAD_USERS = rec
    
End Function

Public Function fn_LOAD_FEATURES(lngUserID As Long) As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    db.strQueryString = "SELECT FeatureID FROM tb_USERS_FEATURES WHERE UserID = " & lngUserID
    Set rec = db.fn_MANAGE_RECORD
    
    If rec.AbsolutePosition <> -1 Then
        mvarFeatureID = rec!FeatureID
    End If
    
    Set fn_LOAD_FEATURES = rec
    
End Function

Public Function fn_LOAD_ALL_FEATURES() As ADODB.Recordset

    Dim rec As New ADODB.Recordset
    db.strQueryString = "SELECT * FROM tb_FEATURES"
    Set rec = db.fn_MANAGE_RECORD
    
    If rec.AbsolutePosition <> -1 Then
        mvarFeatureID = rec!FeatureID
    End If
    
    Set fn_LOAD_ALL_FEATURES = rec
    
End Function

Public Function fn_SAVE_USERS() As ADODB.Recordset

    db.strQueryString = "INSERT INTO tb_USERS " _
    & " (UserID,UserName,FullName,Password,Deleted)" _
    & " VALUES (" _
    & mvarUserID & "," _
    & "'" & mvarUserName & "'," _
    & "'" & mvarFullName & "'," _
    & "'" & mvarPassword & "'," _
    & mvarDeleted & ")"
    
    Set fn_SAVE_USERS = db.fn_MANAGE_RECORD
    
End Function

'Public Function fn_UPDATE_USERS(lngUserID As Long) As ADODB.Recordset
'    Dim Rec As New ADODB.Recordset
'
'    db.strQueryString = "UPDATE tb_Users SET " _
'    & " UserName = '" & mvarUserName & "'," _
'    & " FullName = '" & mvarFullName & "'," _
'    & " Password = '" & mvarPassword & "'" _
'    & " WHERE UserID = " & lngUserID & ""
'
'    Set Rec = db.fn_MANAGE_RECORD
'   Set fn_UPDATE_USERS = Rec
'End Function

Public Function fn_UPDATE_USERS(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_USERS SET " _
    & "UserName = '" & mvarUserName & "'," _
    & "Password = '" & mvarPassword & "'," _
    & "FullName = '" & mvarFullName & "'," _
    & "Status = " & mvarStatus & "," _
    & "Deleted = " & mvarDeleted & "," _
    & "Active = " & mvarActive _
    & " WHERE UserID = " & lngID
    
    Set fn_UPDATE_USERS = db.fn_MANAGE_RECORD

End Function

Public Function fn_DELETE_FEATURES(Optional lngUserID As Long) As ADODB.Recordset

    db.strQueryString = " DELETE FROM tb_USERS_FEATURES WHERE UserID = " & lngUserID

    Set fn_DELETE_FEATURES = db.fn_MANAGE_RECORD

End Function

Public Function fn_SAVE_USERS_FEATURES(lngUserID As Long, lngFeatureID As Long)

    db.strQueryString = "INSERT INTO tb_USERS_FEATURES " _
    & " (UsersFeaturesID,UserID,FeatureID)" _
    & " VALUES (" _
    & mvarUsersFeaturesID & "," _
    & lngUserID & "," _
    & lngFeatureID & ")"
    
    fn_SAVE_USERS_FEATURES = db.fn_MANAGE_RECORD
    
End Function

Public Function fn_DELETE_USERS(Optional lngID As Long) As ADODB.Recordset

    db.strQueryString = "UPDATE tb_USERS SET " _
    & "Deleted = 1 " _
    & " WHERE UserID = " & lngID
    
    Set fn_DELETE_USERS = db.fn_MANAGE_RECORD

End Function

Public Function fn_UPDATE_USER_PASSWORD(lngUserID As Long) As ADODB.Recordset
    db.strQueryString = "UPDATE tb_USERS SET " _
    & " UserName = '" & mvarUserName & "'," _
    & " FullName = '" & mvarFullName & "'," _
    & " Password = '" & mvarPassword & "'" _
    & " WHERE UserID = " & lngUserID & ""
   Set fn_UPDATE_USER_PASSWORD = db.fn_MANAGE_RECORD
End Function

'Public Function fn_UPDATE_USER_NAME(lngUserID As Long) As ADODB.Recordset
'    db.strQueryString = "UPDATE tb_USERS SET " _
'    & " UserName = '" & mvarUserName & "'," _
'    & " FullName = '" & mvarFullName & "'" _
'    & " WHERE UserID = " & lngUserID & ""
'    Set fn_UPDATE_USER_NAME = db.fn_MANAGE_RECORD
'End Function

Public Function fn_AUTOGEN() As Long

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(UserID) FROM tb_USERS"
    
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

Public Function fn_AUTOGEN_USERS_FEATURES_ID() As Long

    Dim rec As New Recordset
    
    db.strQueryString = "SELECT MAX(UsersFeaturesID) FROM tb_USERS_FEATURES"
    
    Set rec = db.fn_MANAGE_RECORD
    
    If rec.AbsolutePosition <> -1 Then
        If IsNull(rec(0)) Then
            fn_AUTOGEN_USERS_FEATURES_ID = 1
        Else
            fn_AUTOGEN_USERS_FEATURES_ID = Val(rec(0)) + 1
        End If
    Else
        fn_AUTOGEN_USERS_FEATURES_ID = 1
    End If

End Function
