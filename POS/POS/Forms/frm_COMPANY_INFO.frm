VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_COMPANY_INFO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COMPANY INFO"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   Icon            =   "frm_COMPANY_INFO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin OCX.b8Container b8Container1 
      Height          =   3225
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   5689
      BackColor       =   16185592
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3075
         Left            =   90
         TabIndex        =   13
         Top             =   60
         Width           =   6045
         Begin VB.TextBox TxtName 
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
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   0
            Top             =   210
            Width           =   4605
         End
         Begin VB.TextBox TxtAddress 
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
            Left            =   1410
            MaxLength       =   100
            TabIndex        =   1
            Top             =   570
            Width           =   4605
         End
         Begin VB.TextBox TxtEmail 
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
            Left            =   1410
            MaxLength       =   100
            TabIndex        =   2
            Top             =   930
            Width           =   4605
         End
         Begin VB.TextBox TxtTelephone 
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
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   3
            Top             =   1290
            Width           =   4605
         End
         Begin VB.TextBox TxtFax 
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
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1650
            Width           =   4605
         End
         Begin VB.TextBox TxtLocation 
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
            Left            =   1410
            MaxLength       =   100
            TabIndex        =   5
            Top             =   2010
            Width           =   4605
         End
         Begin VB.TextBox TxtVATNo 
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
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   6
            Top             =   2370
            Width           =   4605
         End
         Begin VB.TextBox TxtVATRate 
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
            Left            =   1410
            MaxLength       =   6
            TabIndex        =   7
            Top             =   2730
            Width           =   1245
         End
         Begin VB.TextBox TxtNHILRate 
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
            Left            =   4770
            MaxLength       =   6
            TabIndex        =   8
            Top             =   2730
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            TabIndex        =   22
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
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
            TabIndex        =   21
            Top             =   630
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
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
            TabIndex        =   20
            Top             =   990
            Width           =   465
         End
         Begin VB.Label axax 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone No."
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
            TabIndex        =   19
            Top             =   1350
            Width           =   1275
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax No."
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
            TabIndex        =   18
            Top             =   1710
            Width           =   675
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Location"
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
            TabIndex        =   17
            Top             =   2070
            Width           =   750
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VAT No"
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
            TabIndex        =   16
            Top             =   2430
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VAT Rate"
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
            TabIndex        =   15
            Top             =   2790
            Width           =   840
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NHIL Rate"
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
            Left            =   3690
            TabIndex        =   14
            Top             =   2790
            Width           =   915
         End
      End
   End
   Begin OCX.b8Container b8Container2 
      Height          =   735
      Left            =   30
      TabIndex        =   9
      Top             =   3240
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1296
      BackColor       =   16185592
      Begin lvButton.lvButtons_H cmdCancel 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   4590
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
         Image           =   "frm_COMPANY_INFO.frx":000C
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_COMPANY_INFO.frx":057B
      End
      Begin lvButton.lvButtons_H cmdSave 
         Default         =   -1  'True
         Height          =   495
         Left            =   3000
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
         Image           =   "frm_COMPANY_INFO.frx":0895
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_COMPANY_INFO.frx":0AF5
      End
   End
End
Attribute VB_Name = "frm_COMPANY_INFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Mode As String
Dim IsDirty As Boolean
Dim lngID As Long

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub


Private Sub cmdSave_Click()

    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(TxtName, "Please enter company name.") Then Exit Sub
'    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_COMBO(cboContactTitle, "Please select title.") Then Exit Sub
'    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtFirstName, "Please enter first name.") Then Exit Sub
'    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtLastName, "Please enter last name.") Then Exit Sub
    
    With cls_COMPANY_INFO_Obj
        .CompanyName = TxtName.Text
        .Address = txtAddress.Text
        .PhoneNo = TxtTelephone.Text
        .Fax = TxtFax.Text
        .EMail = txtEMail.Text
        .Location = TxtLocation.Text
        .VATNO = TxtVATNo.Text
        .VATRate = TxtVATRate.Text
        .NHILRate = TxtNHILRate.Text
        
        If lngID = 0 Then
            .fn_SAVE_COMPANY_INFO
            Else
                .fn_UPDATE_COMPANY_INFO (lngID)
        End If
        
    End With
    
    MsgBox "Company Profile Saved Successfully", vbInformation, Title
    
End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_CENTER_FORM(Me)
    Call sub_LOAD_COMPANY_INFO
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSave, cmdSave.hwnd, "Save company details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancel, cmdCancel.hwnd, "Cancel and close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
End Sub

Private Sub sub_LOAD_COMPANY_INFO()

    Dim rec As New ADODB.Recordset
    
    Set rec = cls_COMPANY_INFO_Obj.fn_LOAD_COMPANY
    If rec.AbsolutePosition <> -1 Then
        lngID = rec!CompanyID
        TxtName.Text = Trim(rec!CompanyName) & ""
        txtAddress.Text = Trim(rec!Address) & ""
        TxtTelephone.Text = Trim(rec!PhoneNo) & ""
        TxtFax.Text = Trim(rec!Fax) & ""
        txtEMail.Text = Trim(rec!EMail) & ""
        TxtLocation.Text = Trim(rec!Location) & ""
        TxtVATNo.Text = Trim(rec!VATNO) & ""
        TxtVATRate.Text = Trim(rec!VATRate) & ""
        TxtNHILRate.Text = Trim(rec!NHILRate) & ""
    End If
    
    If lngID = 0 Then
        cmdSave.Caption = "Save"
        Else
            cmdSave.Caption = "Update"
    End If
    
End Sub

Private Sub TxtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    IsDirty = True
End Sub


Private Sub TxtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
    IsDirty = True
End Sub


Private Sub TxtFax_KeyDown(KeyCode As Integer, Shift As Integer)
    IsDirty = True
End Sub



Private Sub TxtNHILRate_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub


Private Sub TxtTelephone_KeyDown(KeyCode As Integer, Shift As Integer)
    IsDirty = True
End Sub

Private Sub TxtVATRate_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub
