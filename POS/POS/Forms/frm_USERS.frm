VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_USERS 
   BackColor       =   &H00FFFFFF&
   Caption         =   "USERS"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   12180
   Begin OCX.b8Container b8Container4 
      Height          =   7815
      Left            =   3090
      TabIndex        =   4
      Top             =   0
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   13785
      BackColor       =   16185592
      Begin OCX.b8Container features 
         Height          =   6375
         Left            =   180
         TabIndex        =   5
         Top             =   1350
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   11245
         BackColor       =   16185592
         Begin VB.Frame fraFeatures 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   6195
            Left            =   90
            TabIndex        =   6
            Top             =   90
            Width           =   8325
            Begin VB.CheckBox chkAll 
               Appearance      =   0  'Flat
               BackColor       =   &H00F6F8F8&
               Caption         =   "SELECT ALL FEATURES"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   60
               TabIndex        =   7
               Top             =   60
               Width           =   2835
            End
            Begin MSComctlLib.ListView lvw 
               Height          =   5745
               Left            =   30
               TabIndex        =   8
               Top             =   420
               Width           =   8235
               _ExtentX        =   14526
               _ExtentY        =   10134
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "FeatureName"
                  Object.Width           =   14464
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "FeatureID"
                  Object.Width           =   0
               EndProperty
            End
         End
      End
      Begin OCX.b8Container b8Container5 
         Height          =   1245
         Left            =   180
         TabIndex        =   9
         Top             =   90
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   2196
         BackColor       =   16185592
         Begin VB.Frame fra 
            Appearance      =   0  'Flat
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   90
            TabIndex        =   10
            Top             =   90
            Width           =   8325
            Begin VB.TextBox txtConfirmPassword 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               IMEMode         =   3  'DISABLE
               Left            =   5250
               MaxLength       =   20
               PasswordChar    =   "*"
               TabIndex        =   3
               Top             =   630
               Width           =   2985
            End
            Begin VB.TextBox txtPassword 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               IMEMode         =   3  'DISABLE
               Left            =   5250
               MaxLength       =   20
               PasswordChar    =   "*"
               TabIndex        =   2
               Top             =   120
               Width           =   2985
            End
            Begin VB.TextBox txtFullName 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   1
               Top             =   600
               Width           =   2985
            End
            Begin VB.TextBox txtUserName 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1080
               MaxLength       =   20
               TabIndex        =   0
               Top             =   120
               Width           =   2985
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "User Name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   60
               TabIndex        =   14
               Top             =   150
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Full Name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               TabIndex        =   13
               Top             =   630
               Width           =   1095
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Password"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4350
               TabIndex        =   12
               Top             =   150
               Width           =   1215
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Confirm Password"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   4350
               TabIndex        =   11
               Top             =   630
               Width           =   855
            End
         End
      End
   End
   Begin OCX.b8Container b8Container2 
      Height          =   7815
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   13785
      BackColor       =   16185592
      Begin VB.ListBox lstUsers 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7245
         Left            =   120
         TabIndex        =   16
         Top             =   450
         Width           =   2865
      End
      Begin OCX.b8SideTab b8SideTab2 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   661
         Caption         =   "Users"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   16777215
      End
   End
   Begin OCX.b8Container b8Container3 
      Height          =   675
      Left            =   0
      TabIndex        =   18
      Top             =   7830
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   1191
      BackColor       =   16185592
      Begin lvButton.lvButtons_H cmdAddNew 
         Height          =   495
         Left            =   180
         TabIndex        =   19
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "&Add New"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_USERS.frx":0000
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_USERS.frx":0515
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   8268
         TabIndex        =   20
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "&Cancel"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_USERS.frx":082F
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_USERS.frx":0D9E
      End
      Begin lvButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   2202
         TabIndex        =   21
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "&Edit"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_USERS.frx":10B8
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_USERS.frx":1435
      End
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   495
         Left            =   10290
         TabIndex        =   22
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "Cl&ose"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_USERS.frx":174F
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_USERS.frx":1CDF
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   6246
         TabIndex        =   23
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "&Save"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_USERS.frx":1FF9
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_USERS.frx":2259
      End
      Begin lvButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   4224
         TabIndex        =   24
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "&Delete"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_USERS.frx":2573
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_USERS.frx":2A86
      End
   End
End
Attribute VB_Name = "frm_USERS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim lst As ListItem
Dim blnLogOff As Boolean
Dim blnEdit As Boolean
Dim blnNew As Boolean
Dim lngID As Long

Dim lngNewUserID As Long



Private Sub chkAll_Click()
    With chkAll
        If .Value = 1 Then
            Call Mdl_FUNCTIONS.fn_SELECT_ALL_IN_VIEW(lvw)
            .Caption = "UNSELECT ALL FEATURES"
            Else
                Call Mdl_FUNCTIONS.fn_UNSELECT_ALL_IN_VIEW(lvw)
                .Caption = "SELECT ALL FEATURES"
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    disableCtl
    fra.Enabled = False
    fraFeatures.Enabled = False
    Call sub_LOAD_USERS
    Call InitializeView
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Are you sure you want to delete  " & Trim(txtUserName.Text) & " ?", vbQuestion + vbYesNo, Title) = vbNo Then
        Exit Sub
        Else
            Call cls_USER_Obj.fn_DELETE_USERS(lngID)
            MsgBox "Users successfully deleted.", vbInformation, Title
    End If
    Call sub_CLEAR_FIELD
    Call sub_LOAD_USERS
End Sub

Private Sub cmdEdit_Click()

    If lngUserID <> lstUsers.ItemData(lstUsers.ListIndex) And lngAdminID <> 1 Then
        MsgBox "You don't have the right to edit the details of this user", vbInformation, "Edit User"
        Exit Sub
    End If

'    If lngAdminID <> 1 Then
'        fraFeatures.Enabled = False
'        Else
'            fraFeatures.Enabled = True
'    End If

    fraFeatures.Enabled = True

    blnEdit = True
    blnNew = False
    Call enableCtl
    fra.Enabled = True
    
    txtUserName.SetFocus
    
End Sub

Private Sub cmdaddNew_Click()

    If lngAdminID <> 1 Then
        MsgBox "You don't have the right to create new user", vbInformation, "New User"
        Exit Sub
    End If
    
    lngNewUserID = 0
    blnEdit = False
    blnNew = True
    Call enableCtl
    fra.Enabled = True
    fraFeatures.Enabled = True
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call InitializeView
    txtUserName.SetFocus
End Sub

Private Sub cmdSave_Click()
On Error GoTo errHandler
    Dim ctr As Long
    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtUserName, "Please enter the user name") = True Then Exit Sub
    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtFullName, "Please enter the full name") = True Then Exit Sub
    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtPassword, "Please enter the password") = True Then Exit Sub
    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtConfirmPassword, "Please confirm password") = True Then Exit Sub

    If Trim(txtPassword.Text) <> Trim(txtConfirmPassword.Text) Then
        MsgBox "The passwords do not match", vbInformation, "Save"
        Exit Sub
    End If

    With cls_USER_Obj
        
        .UserName = Trim(txtUserName.Text)
        .Password = Trim(txtPassword.Text)
        .FullName = Trim(txtFullName.Text)
        .UsersFeaturesID = .fn_AUTOGEN_USERS_FEATURES_ID

        If lngCurrentUserID = lngID And MsgBox("You will have to log on again. Do you want to continue?", vbQuestion + vbYesNo, "User") = vbNo Then Exit Sub
            
            If blnNew = True Then
                    .UserID = .fn_AUTOGEN
                    Call .fn_SAVE_USERS
                Else
                    Call .fn_UPDATE_USERS(lngID)
                    Call .fn_DELETE_FEATURES(lstUsers.ItemData(lstUsers.ListIndex))
                    For ctr = 1 To lvw.ListItems.Count
                        If lvw.ListItems(ctr).Checked = True Then
                            Call .fn_SAVE_USERS_FEATURES(lstUsers.ItemData(lstUsers.ListIndex), lvw.ListItems(ctr).ListSubItems(1).Text)
                        End If
                    Next
            End If
            
            MsgBox "Changes made successfully", vbExclamation, Title
            Call disableCtl
            Call sub_CLEAR_FIELD
            Call sub_LOAD_USERS
            fraFeatures.Enabled = False
            
        End With

'    Call frm_MAIN.sub_DISABLE_FEATURES
'    Call frmMain.mnuCloseAll_Click
'    frmLogin.Show 1
    
EXITPROCEDURE:
    Exit Sub

errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE

End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_FORM_SIZE(Me)
    Move 0, 0
    disableCtl
    Call sub_LOAD_USERS
    Call sub_LOAD_ALL_FEATURES
    
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdAddNew, cmdAddNew.hwnd, "Add new user details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdEdit, cmdEdit.hwnd, "Edit existing user details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdDelete, cmdDelete.hwnd, "Delete existing user details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSave, cmdSave.hwnd, "Save user details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancel, cmdCancel.hwnd, "Cancel the transaction.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdClose, cmdClose.hwnd, "Close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
   
   
End Sub

Private Sub sub_LOAD_USERS()

    Dim rec As New ADODB.Recordset

    Set rec = cls_USER_Obj.fn_LOAD_USERS(0)

    lstUsers.Clear

    If rec.AbsolutePosition = -1 Then Exit Sub
    Do While Not rec.EOF
        lstUsers.AddItem rec!UserName
        lstUsers.ItemData(lstUsers.NewIndex) = rec!UserID
        rec.MoveNext
    Loop

End Sub

Private Sub sub_LOAD_USERS_DETAILS(lngUserID As Long)

    Dim rec As New ADODB.Recordset

    Set rec = cls_USER_Obj.fn_LOAD_USERS(lngUserID)

    If rec.AbsolutePosition = -1 Then Exit Sub
    txtUserName.Text = Trim(rec!UserName)
    txtFullName.Text = Trim(rec!FullName)
    txtPassword.Text = Trim(rec!Password)
    txtConfirmPassword.Text = Trim(rec!Password)
    
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    blnEdit = False
    blnNew = False
End Sub

Private Sub lstUsers_Click()
    Call InitializeView
    lngID = lstUsers.ItemData(lstUsers.ListIndex)
    Call sub_LOAD_USERS_DETAILS(lngID)
    Call sub_LOAD_USERS_FEATURES(lngID)
End Sub

Private Sub sub_LOAD_ALL_FEATURES()

    Dim rec As New ADODB.Recordset

    Set rec = cls_USER_Obj.fn_LOAD_ALL_FEATURES

    If rec.AbsolutePosition = -1 Then Exit Sub

    Do While Not rec.EOF
        Set lst = lvw.ListItems.Add(, , rec!FeatureName & "")
            lst.ListSubItems.Add , , rec!FeatureID
        rec.MoveNext
    Loop


End Sub

Private Sub sub_LOAD_USERS_FEATURES(lngUserID As Long)

    Dim rec As New ADODB.Recordset
    Dim ctr As Long
    Set rec = cls_USER_Obj.fn_LOAD_FEATURES(lngUserID)

    If rec.AbsolutePosition = -1 Then Exit Sub

    Do While Not rec.EOF
        For ctr = 1 To lvw.ListItems.Count
            If lvw.ListItems(ctr).ListSubItems(1).Text = rec!FeatureID Then
                lvw.ListItems(ctr).Checked = True
            End If
        Next
        rec.MoveNext
    Loop

End Sub

Private Sub disableCtl()

    cmdAddNew.Enabled = True
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = False
    cmdEdit.Enabled = False

End Sub

Private Sub enableCtl()

    cmdAddNew.Enabled = False
    cmdSave.Enabled = True
    cmdDelete.Enabled = True
    cmdCancel.Enabled = True
    cmdEdit.Enabled = True

End Sub

Private Sub InitializeView()

    Dim ctr As Long
    For ctr = 1 To lvw.ListItems.Count
        lvw.ListItems(ctr).Checked = False
    Next

End Sub

Private Sub sub_CLEAR_FIELD()

    txtUserName.Text = ""
    txtFullName.Text = ""
    txtPassword.Text = ""
    txtConfirmPassword.Text = ""

End Sub

