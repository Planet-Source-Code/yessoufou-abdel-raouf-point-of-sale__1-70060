VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_USER_NAME 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CHANGE USER NAME"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00F6F8F8&
      Height          =   2325
      Left            =   0
      TabIndex        =   5
      Top             =   -30
      Width           =   5325
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
         Height          =   315
         Left            =   1710
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   3435
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
         Height          =   315
         Left            =   1710
         MaxLength       =   100
         TabIndex        =   1
         Top             =   660
         Width           =   3435
      End
      Begin VB.TextBox txtOldPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         MaxLength       =   20
         PasswordChar    =   "v"
         TabIndex        =   2
         Top             =   1080
         Width           =   3435
      End
      Begin VB.TextBox txtNewPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         MaxLength       =   16
         PasswordChar    =   "v"
         TabIndex        =   3
         Top             =   1470
         Width           =   3435
      End
      Begin VB.TextBox txtConfirmPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         MaxLength       =   16
         PasswordChar    =   "v"
         TabIndex        =   4
         Top             =   1845
         Width           =   3435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   300
         Width           =   945
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1470
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1860
         Width           =   1515
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1110
         Width           =   1170
      End
   End
   Begin OCX.b8Container b8Container2 
      Height          =   735
      Left            =   0
      TabIndex        =   11
      Top             =   2310
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   1296
      BackColor       =   16185592
      Begin lvButton.lvButtons_H cmdCancel 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   3600
         TabIndex        =   12
         Top             =   120
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
         Image           =   "frm_USER_NAME.frx":0000
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_USER_NAME.frx":056F
      End
      Begin lvButton.lvButtons_H cmdSave 
         Default         =   -1  'True
         Height          =   495
         Left            =   1740
         TabIndex        =   13
         Top             =   120
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
         Image           =   "frm_USER_NAME.frx":0889
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_USER_NAME.frx":0AE9
      End
   End
End
Attribute VB_Name = "frm_USER_NAME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo errHandler

    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtUserName, "Kindly enter New User Name.") Then GoTo EXITPROCEDURE
    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtFullName, "Kindly confirm the Full Name.") Then GoTo EXITPROCEDURE
    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtOldPassword, "Kindly enter Old Password.") Then GoTo EXITPROCEDURE
    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtNewPassword, "Kindly confirm the New Password.") Then GoTo EXITPROCEDURE
    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtConfirmPassword, "Kindly Confirm New Password.") Then GoTo EXITPROCEDURE

    If txtNewPassword.Text <> txtConfirmPassword.Text Then
        MsgBox "The new Password are not the same", vbExclamation, "Change Password"
        Exit Sub
    End If

    With cls_USER_Obj
        If .fn_CHECK_USER_LOGIN(strUserName, Trim(txtOldPassword)) Then
    
                If MsgBox("This will stop all transactions and take you to the login." & vbCrLf & " Do you really want to save?", vbQuestion + vbYesNo, "User Details") = vbNo Then Exit Sub
    
                .UserName = Trim(txtUserName.Text)
                .FullName = Trim(txtFullName.Text)
                .Password = Trim(txtNewPassword.Text)
                Call .fn_UPDATE_USER_PASSWORD(lngCurrentUserID)
                Call frm_LOGIN.sub_SAVE_USERS_LOGS(0, "Log Off")
                Call Mdl_FUNCTIONS.sub_CLOSE_ALL_OPENED_FORMS
                frm_MAIN.Show
                frm_LOGIN.Show 1
    
            Else
    
                MsgBox "Invalid Old Password.", vbExclamation, "User name"
                Call Mdl_FUNCTIONS.fn_HIGHLIGHT_TEXT(txtOldPassword)
    
        End If
    End With
EXITPROCEDURE:
    Exit Sub

errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub


Private Sub Form_Load()

    Call Mdl_FUNCTIONS.sub_CENTER_FORM(Me)
    txtUserName.Text = strUserName
    txtFullName.Text = strFullName

    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSave, cmdSave.hwnd, "Save user details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancel, cmdCancel.hwnd, "Cancel and close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    
End Sub

Private Sub txtFullName_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_ALPHABET_ONLY(KeyAscii)
End Sub
