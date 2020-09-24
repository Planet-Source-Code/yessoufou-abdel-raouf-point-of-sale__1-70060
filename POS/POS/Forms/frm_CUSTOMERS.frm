VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_CUSTOMERS 
   Caption         =   "CUSTOMERS"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   12180
   Begin OCX.b8Container b8Container1 
      Height          =   7905
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   13944
      BackColor       =   16185592
      Begin VB.ListBox lst 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7635
         Left            =   150
         TabIndex        =   10
         Top             =   150
         Width           =   2715
      End
      Begin VB.Frame fraDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   7725
         Left            =   2970
         TabIndex        =   11
         Top             =   60
         Width           =   6195
         Begin VB.ComboBox cboContactTitle 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   750
            Width           =   1935
         End
         Begin VB.TextBox txtAddress 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1245
            Left            =   1695
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   8
            Text            =   "frm_CUSTOMERS.frx":0000
            Top             =   4050
            Width           =   3855
         End
         Begin VB.TextBox txtEMail 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1695
            MaxLength       =   100
            TabIndex        =   7
            Text            =   " "
            Top             =   3570
            Width           =   3855
         End
         Begin VB.TextBox txtPhoneNo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1695
            MaxLength       =   20
            TabIndex        =   6
            Top             =   3105
            Width           =   3885
         End
         Begin VB.ComboBox cboGender 
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
            ItemData        =   "frm_CUSTOMERS.frx":0002
            Left            =   1695
            List            =   "frm_CUSTOMERS.frx":000C
            TabIndex        =   5
            Top             =   2640
            Width           =   1965
         End
         Begin VB.TextBox txtOtherNames 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1695
            MaxLength       =   100
            TabIndex        =   4
            Text            =   " "
            Top             =   2172
            Width           =   3855
         End
         Begin VB.TextBox txtLastName 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1695
            MaxLength       =   50
            TabIndex        =   3
            Text            =   " "
            Top             =   1704
            Width           =   3855
         End
         Begin VB.TextBox txtFirstName 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1695
            MaxLength       =   50
            TabIndex        =   2
            Text            =   " "
            Top             =   1236
            Width           =   3885
         End
         Begin VB.TextBox txtCustomerNo 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1695
            MaxLength       =   20
            TabIndex        =   0
            Text            =   " "
            Top             =   300
            Width           =   3885
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   270
            TabIndex        =   20
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "First Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   270
            TabIndex        =   19
            Top             =   1230
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Othe Names"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   270
            TabIndex        =   18
            Top             =   2175
            Width           =   1485
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   270
            TabIndex        =   17
            Top             =   1710
            Width           =   1455
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Title"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   270
            TabIndex        =   16
            Top             =   765
            Width           =   1305
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   270
            TabIndex        =   15
            Top             =   2640
            Width           =   1485
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   270
            TabIndex        =   14
            Top             =   3105
            Width           =   1215
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   270
            TabIndex        =   13
            Top             =   4050
            Width           =   1485
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "EMail"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   270
            TabIndex        =   12
            Top             =   3570
            Width           =   1455
         End
      End
      Begin OCX.b8Container fra 
         Height          =   7665
         Left            =   9300
         TabIndex        =   21
         Top             =   150
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   13520
         BackColor       =   16185592
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2685
            Left            =   150
            ScaleHeight     =   2655
            ScaleWidth      =   2355
            TabIndex        =   22
            Top             =   210
            Width           =   2385
            Begin VB.Image imgPicture 
               Height          =   2655
               Left            =   0
               Picture         =   "frm_CUSTOMERS.frx":001E
               Stretch         =   -1  'True
               Top             =   0
               Width           =   2355
            End
            Begin VB.Label lblAlerte 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "NO PICTURE AVAILABLE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   90
               TabIndex        =   23
               Top             =   1110
               Width           =   1965
            End
         End
         Begin lvButton.lvButtons_H cmdBrowse 
            Height          =   435
            Left            =   150
            TabIndex        =   24
            Top             =   3390
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   767
            Caption         =   "&Browse..."
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
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_CUSTOMERS.frx":26D6
         End
         Begin MSComDlg.CommonDialog PictureDlg 
            Left            =   240
            Top             =   4620
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin lvButton.lvButtons_H cmdIdentification 
            Height          =   375
            Left            =   150
            TabIndex        =   25
            Top             =   2880
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   661
            CapAlign        =   2
            BackStyle       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   16777215
            cFHover         =   16777215
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   16711680
         End
         Begin lvButton.lvButtons_H cmdRemove 
            Height          =   435
            Left            =   150
            TabIndex        =   26
            Top             =   3960
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   767
            Caption         =   "&Remove"
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
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_CUSTOMERS.frx":29F0
         End
      End
   End
   Begin OCX.b8Container b8Container3 
      Height          =   675
      Left            =   0
      TabIndex        =   27
      Top             =   7920
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   1191
      BackColor       =   16185592
      Begin lvButton.lvButtons_H cmdAddNew 
         Height          =   495
         Left            =   180
         TabIndex        =   28
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
         Image           =   "frm_CUSTOMERS.frx":2D0A
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_CUSTOMERS.frx":321F
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   8436
         TabIndex        =   29
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
         Image           =   "frm_CUSTOMERS.frx":3539
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_CUSTOMERS.frx":3AA8
      End
      Begin lvButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   2244
         TabIndex        =   30
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
         Image           =   "frm_CUSTOMERS.frx":3DC2
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_CUSTOMERS.frx":413F
      End
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   495
         Left            =   10500
         TabIndex        =   31
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
         Image           =   "frm_CUSTOMERS.frx":4459
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_CUSTOMERS.frx":49E9
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   6372
         TabIndex        =   32
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
         Image           =   "frm_CUSTOMERS.frx":4D03
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_CUSTOMERS.frx":4F63
      End
      Begin lvButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   4308
         TabIndex        =   33
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
         Image           =   "frm_CUSTOMERS.frx":527D
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_CUSTOMERS.frx":5790
      End
   End
End
Attribute VB_Name = "frm_CUSTOMERS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngID As Long
Dim blnEdit As Boolean
Dim blnAdd As Boolean

Private Sub cboGender_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo errHandler
    
'    imgPicture.Picture = LoadPicture()
    PictureDlg.ShowOpen
    If PictureDlg.FileName = "" Then GoTo EXITPROCEDURE
    imgPicture.Picture = LoadPicture(PictureDlg.FileName)

EXITPROCEDURE:
    Exit Sub
    
errHandler:
    MsgBox "Error loading picture", vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, Me.Name, "cmdTelecharger")
    GoTo EXITPROCEDURE
End Sub

Private Sub cmdRemove_Click()
On Error GoTo errHandler

    If imgPicture.Picture = False Then Exit Sub

    If MsgBox("Are you sure you want to remove  " & Trim(txtFirstName.Text) & "'s picture ?", vbQuestion + vbYesNo, Title) = vbNo Then
        
        GoTo EXITPROCEDURE
        
        Else
        
            Call mdl_GLOBAL_VARIABLES.cls_CUSTOMER_Obj.fn_DELETE_PICTURE(lngID)
            strPictureName = Trim(cmdIdentification.Caption) & ".bmp"
            Kill strPicturePath & "\Pictures\" & strPictureName
            imgPicture.Picture = LoadPicture()
            
            Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
            Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
            Call subDisableCtl
            Call subLoadCustomers
    End If
    
EXITPROCEDURE:
    Exit Sub

errHandler:
    MsgBox "Error loading picture", vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, Me.Name, "cmdTelecharger")
    GoTo EXITPROCEDURE
End Sub


Private Sub cmdaddNew_Click()

    blnAdd = True
    blnEdit = False
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call Mdl_FUNCTIONS.fn_UNSET_CONTROL_COLOR(Me, "All")
    cmdIdentification.Caption = cls_CUSTOMER_Obj.fn_AUTOGEN
    txtCustomerNo.Text = cls_CUSTOMER_Obj.fn_AUTOGEN
    Call subEnableCtl
    cboContactTitle.SetFocus
    
End Sub

Private Sub cmdCancel_Click()
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    Call subDisableCtl
    Call subLoadCustomers
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo errHandler
    If MsgBox("Are you sure you want to Delete Customer " & Trim(txtCustomerNo.Text) & "?", vbQuestion + vbYesNo, Title) = vbNo Then Exit Sub

    With cls_CUSTOMER_Obj
        .fn_CHECK_CUSTOMER_IN_SALES (lngID)
        If blnCustomerExist = True Then
            MsgBox "Customer Can Not Be Deleted.", vbInformation, Title
            Exit Sub
            Else
                Call cls_CUSTOMER_Obj.fn_DELETE_CUSTOMER(lngID)
                Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
                Call subLoadCustomers
        End If
    End With

EXITPROCEDURE:
    Exit Sub

errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub

Private Sub cmdEdit_Click()
    blnAdd = False
    blnEdit = True
    Call Mdl_FUNCTIONS.fn_UNSET_CONTROL_COLOR(Me, "All")
    Call subEnableCtl
End Sub

Private Sub cmdSave_Click()
On Error GoTo errHandler

    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtCustomerNo, "Please enter passenger ID.") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_COMBO(cboContactTitle, "Please select title.") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtFirstName, "Please enter first name.") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtLastName, "Please enter last name.") Then Exit Sub
    
    'If Mdl_FUNCTIONS.fn_CHECK_EMPTY_COMBO(cboGender, "Please select gender.") Then Exit Sub
    If cboGender.Text = "" Then
        MsgBox "Please select the gender", vbExclamation, "Gender"
        cboGender.SetFocus
        Exit Sub
    End If
    
    'If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtFinalDestination, "Please enter  final destination.") Then Exit Sub

    With cls_CUSTOMER_Obj
        .CustomerID = .fn_ID_AUTOGEN
        .CustomerNo = Trim(cmdIdentification.Caption)
        .FirstName = Trim(txtFirstName.Text)
        .LastName = Trim(txtLastName.Text)
        .OtherNames = Trim(txtOtherNames.Text)
        .Title = cboContactTitle.ItemData(cboContactTitle.ListIndex)
        .Gender = Trim(cboGender.Text)
        .PhoneNo = Trim(txtPhoneNo.Text)
        .EMail = Trim(txtEMail.Text)
        .Address = Trim(txtAddress.Text)
        strPicturePath = App.Path
        If imgPicture.Picture Then
            strPictureName = Trim(cmdIdentification.Caption) & ".bmp"
            Call SavePicture(imgPicture, strPicturePath & "\Pictures\" & strPictureName)
            .Picture = strPictureName
            Else
                .Picture = ""
        End If
        
        If blnAdd = True And blnEdit = False Then
            Call .fn_SAVE_CUSTOMER_RECORDS
            Else
                Call .fn_UPDATE_CUSTOMER_RECORDS(lngID)
        End If
    End With

    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    Call subDisableCtl
    Call subLoadCustomers

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
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call sub_LOAD_TITLES(cboContactTitle)
    Call subLoadCustomers
    Call subDisableCtl

    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdAddNew, cmdAddNew.hwnd, "Add new customer details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdEdit, cmdEdit.hwnd, "Edit existing customer details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdDelete, cmdDelete.hwnd, "Delete existing customer details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSave, cmdSave.hwnd, "Save customer details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancel, cmdCancel.hwnd, "Cancel the transaction.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdClose, cmdClose.hwnd, "Close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
   
   
End Sub

Private Sub sub_LOAD_TITLES(Optional cbo As ComboBox)

    Dim rec As New ADODB.Recordset
    Set rec = cls_REFERENCES_Obj.fn_LOADT_TITLES
    
    cbo.Clear
    
    Do While Not rec.EOF
        cbo.AddItem rec!TitleName
        cbo.ItemData(cbo.NewIndex) = rec!TitleID
        rec.MoveNext
    Loop

End Sub

Private Sub subLoadCustomers()

    Dim rec As New ADODB.Recordset
    Set rec = cls_CUSTOMER_Obj.fn_LOAD_CUSTOMERS(0)

    lst.Clear

    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            lst.AddItem rec!FirstName & " " & rec!LastName
            lst.ItemData(lst.NewIndex) = rec!CustomerID
            rec.MoveNext
        Loop
    End If

End Sub

Private Sub subLoadCustomersDetails(lngID As Long)
On Error GoTo errHandler
    Dim rec As New ADODB.Recordset
    Set rec = cls_CUSTOMER_Obj.fn_LOAD_CUSTOMERS(lngID)

    If rec.AbsolutePosition <> -1 Then
        txtCustomerNo.Text = rec!CustomerNo
        cmdIdentification.Caption = rec!CustomerNo
        txtFirstName.Text = rec!FirstName
        txtLastName.Text = rec!LastName
        txtOtherNames.Text = rec!OtherNames
        cboContactTitle.ListIndex = Mdl_FUNCTIONS.fn_GET_LIST_INDEX(cboContactTitle, rec!Title)
        cboGender.Text = rec!Gender
        txtPhoneNo.Text = rec!PhoneNo
        txtEMail.Text = rec!EMail
        txtAddress.Text = rec!Address
        If rec!Picture = "" Then
            imgPicture.Picture = LoadPicture("")
            Else
                imgPicture.Picture = LoadPicture(App.Path & "\Pictures\" & rec!Picture)
        End If
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
    End If
    
EXITPROCEDURE:
    Exit Sub

errHandler:
    If Err.Number = 53 Then
        imgPicture.Picture = LoadPicture()
        Exit Sub
    End If
    MsgBox "Error Occurred while loading picture", vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, Me.Name, "cmdTelecharger")
    
    GoTo EXITPROCEDURE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blnAdd = False
    blnEdit = False
End Sub


Private Sub subDisableCtl()

    cmdAddNew.Enabled = True
    cmdClose.Enabled = True

    cmdBrowse.Enabled = False
    cmdRemove.Enabled = False
    cmdEdit.Enabled = False
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = False

    fraDetails.Enabled = False
    lst.Enabled = True
End Sub

Private Sub subEnableCtl()

    cmdAddNew.Enabled = False
    cmdClose.Enabled = False

    cmdBrowse.Enabled = True
    cmdRemove.Enabled = True
    cmdEdit.Enabled = False
    cmdSave.Enabled = True
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    fraDetails.Enabled = True

    lst.Enabled = False
End Sub

Private Sub lst_Click()
    lngID = lst.ItemData(lst.ListIndex)
    Call subLoadCustomersDetails(lngID)
End Sub



Private Sub lst_DblClick()
    If blnCustomerDeposit = True Then
        With frm_CUSTOMER_DEPOSIT
            .txtCashCustomer.Text = cboContactTitle.Text & " " & txtFirstName.Text & " " & txtLastName.Text
            .txtChequeCustomer.Text = cboContactTitle.Text & " " & txtFirstName.Text & " " & txtLastName.Text
            .txtBankDepositCustomer.Text = cboContactTitle.Text & " " & txtFirstName.Text & " " & txtLastName.Text
            lngSelectedCustomerID = lst.ItemData(lst.ListIndex)
        End With
    End If
    Unload Me
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_ALPHABET_ONLY(KeyAscii)
End Sub



Private Sub txtFirstName_Validate(Cancel As Boolean)
    txtFirstName.Text = StrConv(txtFirstName, vbProperCase)
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_ALPHABET_ONLY(KeyAscii)
End Sub

Private Sub txtLastName_Validate(Cancel As Boolean)
    txtLastName.Text = StrConv(txtLastName, vbProperCase)
End Sub

Private Sub txtOtherNames_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_ALPHABET_ONLY(KeyAscii)
End Sub

Private Sub txtOtherNames_Validate(Cancel As Boolean)
    txtOtherNames.Text = StrConv(txtOtherNames, vbProperCase)
End Sub

Private Sub txtPhoneNo_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_ALPHABET_ONLY(KeyAscii)
End Sub
