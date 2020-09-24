VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_SERVER_CONNECTION 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SERVER CONNECTION"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5130
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_SERVER_CONNECTION.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin OCX.b8Container b8Container2 
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   2190
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   1296
      BackColor       =   16185592
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   3390
         TabIndex        =   10
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
         Image           =   "frm_SERVER_CONNECTION.frx":0442
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SERVER_CONNECTION.frx":09B1
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1560
         TabIndex        =   11
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
         Image           =   "frm_SERVER_CONNECTION.frx":0CCB
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SERVER_CONNECTION.frx":0F2B
      End
   End
   Begin OCX.b8Container b8Container1 
      Height          =   2205
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   3889
      BackColor       =   16185592
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
         IMEMode         =   3  'DISABLE
         Left            =   1530
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   3315
      End
      Begin VB.TextBox txtPassword 
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
         Left            =   1530
         MaxLength       =   20
         PasswordChar    =   "v"
         TabIndex        =   1
         Top             =   720
         Width           =   3315
      End
      Begin VB.TextBox txtServerName 
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
         IMEMode         =   3  'DISABLE
         Left            =   1530
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1170
         Width           =   3315
      End
      Begin VB.TextBox txtDatabase 
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
         IMEMode         =   3  'DISABLE
         Left            =   1530
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1650
         Width           =   3315
      End
      Begin VB.Label Label5 
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
         Height          =   375
         Left            =   180
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
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
         Height          =   375
         Left            =   180
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Server Name"
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
         Left            =   180
         TabIndex        =   7
         Top             =   1170
         Width           =   1155
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Database"
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
         Left            =   180
         TabIndex        =   6
         Top             =   1650
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frm_SERVER_CONNECTION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo errHandler

    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtServerName, "Please enter the Server Name") = True Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtUserName, "Please enter the User Name") = True Then Exit Sub
'    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtPassword, "Please enter the Password Name") = True Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtDatabase, "Please enter the Database Name") = True Then Exit Sub

    If MDL_INI.WriteIni(App.Title, "Server Name", Trim(txtServerName.Text)) = True And _
         MDL_INI.WriteIni(App.Title, "User Name", Trim(txtUserName.Text)) = True And _
             MDL_INI.WriteIni(App.Title, "Password", Trim(txtPassword.Text)) = True And _
                MDL_INI.WriteIni(App.Title, "Database", Trim(txtDatabase.Text)) = True Then
                MsgBox "Settings successfully saved", vbInformation, "Settings"
                Unload Me
    Else
        MsgBox "An error occurred while saving settings. Please retry", vbExclamation, Title
    End If
    

EXITPROCEDURE:
    Exit Sub

errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_CENTER_FORM(Me)
    With SystemData
        txtServerName.Text = .DB_ServerName
        txtUserName.Text = .DB_UserName
        txtPassword.Text = .DB_Password
        txtDatabase.Text = .DB_Database
    End With
End Sub


Public Sub CloseAllOpenedForms()

    Dim i As Integer
    For i = Forms.Count - 1 To 0 Step -1
        If Forms(i).Name <> "frmMain" Then
            Unload Forms(i)
        End If
    Next i

End Sub

