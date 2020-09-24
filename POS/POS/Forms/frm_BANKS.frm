VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_BANKS 
   BackColor       =   &H00FFFFFF&
   Caption         =   "BANKS & ACCOUNTS"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   12180
   Begin VB.Frame Frame1 
      BackColor       =   &H00F6F8F8&
      Height          =   4695
      Left            =   30
      TabIndex        =   8
      Top             =   3780
      Width           =   11955
      Begin VB.ListBox lstBankAccount 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3390
         Left            =   120
         TabIndex        =   16
         Top             =   570
         Width           =   3015
      End
      Begin VB.Frame fra 
         BackColor       =   &H00F6F8F8&
         Enabled         =   0   'False
         Height          =   3495
         Left            =   3240
         TabIndex        =   9
         Top             =   480
         Width           =   8625
         Begin VB.TextBox txtAccountName 
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
            IMEMode         =   3  'DISABLE
            Left            =   1740
            MaxLength       =   50
            TabIndex        =   4
            Top             =   780
            Width           =   3435
         End
         Begin VB.TextBox txtOpeningBalance 
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
            IMEMode         =   3  'DISABLE
            Left            =   1740
            MaxLength       =   20
            TabIndex        =   6
            Top             =   1590
            Width           =   3435
         End
         Begin VB.TextBox txtAccountNo 
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
            Left            =   1740
            MaxLength       =   20
            TabIndex        =   3
            Text            =   " "
            Top             =   360
            Width           =   3435
         End
         Begin VB.ComboBox cboAccountType 
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
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1200
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker DTPOpenDate 
            Height          =   345
            Left            =   1740
            TabIndex        =   7
            Top             =   1980
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   58916865
            CurrentDate     =   39252
         End
         Begin lvButton.lvButtons_H cmdLoadAccountType 
            Height          =   375
            Left            =   3570
            TabIndex        =   10
            Top             =   1140
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   661
            Caption         =   "&Add New ..."
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
            Image           =   "frm_BANKS.frx":0000
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANKS.frx":0515
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Account Type"
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
            Left            =   180
            TabIndex        =   15
            Top             =   1230
            Width           =   1245
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   14
            Top             =   750
            Width           =   1305
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Open Date"
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
            Left            =   180
            TabIndex        =   13
            Top             =   1980
            Width           =   1215
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Account No"
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
            Left            =   180
            TabIndex        =   12
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Opening Balance"
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
            Left            =   180
            TabIndex        =   11
            Top             =   1620
            Width           =   1425
         End
      End
      Begin OCX.b8Container b8Container2 
         Height          =   675
         Left            =   90
         TabIndex        =   17
         Top             =   3960
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   1191
         BackColor       =   16185592
         Begin lvButton.lvButtons_H cmdNewAccount 
            Height          =   495
            Left            =   180
            TabIndex        =   18
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
            Image           =   "frm_BANKS.frx":082F
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANKS.frx":0D44
         End
         Begin lvButton.lvButtons_H cmdCancelAccount 
            Height          =   495
            Left            =   8220
            TabIndex        =   19
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
            Image           =   "frm_BANKS.frx":105E
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANKS.frx":15CD
         End
         Begin lvButton.lvButtons_H cmdEditAccount 
            Height          =   495
            Left            =   2190
            TabIndex        =   20
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
            Image           =   "frm_BANKS.frx":18E7
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANKS.frx":1C64
         End
         Begin lvButton.lvButtons_H cmdCloseAccount 
            Height          =   495
            Left            =   10170
            TabIndex        =   21
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
            Image           =   "frm_BANKS.frx":1F7E
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANKS.frx":250E
         End
         Begin lvButton.lvButtons_H cmdSaveAccount 
            Height          =   495
            Left            =   6210
            TabIndex        =   22
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
            Image           =   "frm_BANKS.frx":2828
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANKS.frx":2A88
         End
         Begin lvButton.lvButtons_H cmdDeleteAccount 
            Height          =   495
            Left            =   4200
            TabIndex        =   23
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
            Image           =   "frm_BANKS.frx":2DA2
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANKS.frx":32B5
         End
      End
      Begin OCX.b8SideTab b8SideTab2 
         Height          =   345
         Left            =   0
         TabIndex        =   24
         Top             =   120
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   609
         Caption         =   "BANK ACCOUNTS"
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
      Height          =   3795
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   6694
      BackColor       =   16185592
      Begin VB.Frame fraBankDetails 
         BackColor       =   &H00F6F8F8&
         Enabled         =   0   'False
         Height          =   2865
         Left            =   3270
         TabIndex        =   27
         Top             =   60
         Width           =   8595
         Begin VB.TextBox txtShortName 
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
            Left            =   1740
            MaxLength       =   100
            TabIndex        =   0
            Text            =   " "
            Top             =   300
            Width           =   3435
         End
         Begin VB.TextBox txtBranch 
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
            Left            =   1740
            MaxLength       =   20
            TabIndex        =   2
            Text            =   " "
            Top             =   1095
            Width           =   3435
         End
         Begin VB.TextBox txtBankName 
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
            Left            =   1740
            MaxLength       =   20
            TabIndex        =   1
            Text            =   " "
            Top             =   705
            Width           =   3435
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Short Name"
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
            Left            =   150
            TabIndex        =   30
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Branch"
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
            Left            =   150
            TabIndex        =   29
            Top             =   1095
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name"
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
            Left            =   150
            TabIndex        =   28
            Top             =   705
            Width           =   1485
         End
      End
      Begin VB.ListBox lstBank 
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
         Height          =   2370
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   3015
      End
      Begin OCX.b8Container b8Container4 
         Height          =   675
         Left            =   90
         TabIndex        =   31
         Top             =   3000
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   1191
         BackColor       =   16185592
         Begin lvButton.lvButtons_H cmdNewBank 
            Height          =   495
            Left            =   180
            TabIndex        =   32
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
            Image           =   "frm_BANKS.frx":35CF
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANKS.frx":3AE4
         End
         Begin lvButton.lvButtons_H cmdCancelBank 
            Height          =   495
            Left            =   8172
            TabIndex        =   33
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
            Image           =   "frm_BANKS.frx":3DFE
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANKS.frx":436D
         End
         Begin lvButton.lvButtons_H cmdEditBank 
            Height          =   495
            Left            =   2178
            TabIndex        =   34
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
            Image           =   "frm_BANKS.frx":4687
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANKS.frx":4A04
         End
         Begin lvButton.lvButtons_H cmdCloseBank 
            Height          =   495
            Left            =   10170
            TabIndex        =   35
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
            Image           =   "frm_BANKS.frx":4D1E
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANKS.frx":52AE
         End
         Begin lvButton.lvButtons_H cmdSaveBank 
            Height          =   495
            Left            =   6174
            TabIndex        =   36
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
            Image           =   "frm_BANKS.frx":55C8
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANKS.frx":5828
         End
         Begin lvButton.lvButtons_H cmdDeleteBank 
            Height          =   495
            Left            =   4176
            TabIndex        =   37
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
            Image           =   "frm_BANKS.frx":5B42
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANKS.frx":6055
         End
      End
      Begin OCX.b8SideTab b8SideTab1 
         Height          =   345
         Left            =   120
         TabIndex        =   38
         Top             =   150
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   609
         Caption         =   "BANKS"
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
End
Attribute VB_Name = "frm_BANKS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim lngID As Long
Dim lngAccountID As Long

Dim blnEdit As Boolean
Dim blnAdd As Boolean

Dim blnEditAccount As Boolean
Dim blnAddAccount As Boolean


Private Sub cmdLoadAccountType_Click()
    blnAccountType = True
    frmAccountType.Show 1
End Sub

Private Sub cmdNewBank_Click()
    blnAdd = True
    blnEdit = False
    Call fn_UNSET_CONTROL_COLOR("Bank")
    Call sub_EMPTY_FIELS("Bank")
    Call subEnableCtl
    txtShortName.SetFocus
End Sub

Private Sub cmdCancelBank_Click()
   Call sub_EMPTY_FIELS("Bank")
    Call fn_SET_CONTROL_COLOR("Bank")
    Call subDisableCtl
    Call subLoadBank
End Sub

Private Sub cmdCloseBank_Click()
    Unload Me
End Sub

Private Sub cmdDeleteBank_Click()
On Error GoTo errHandler

    If MsgBox("Are you sure you want to Delete Agency  " & Trim(txtAgencyName.Text) & "?", vbQuestion + vbYesNo, "Delete Aircraft") = vbNo Then Exit Sub
    
    Call cls_AGENCY_Obj.fn_DELETE_AGENCY_RECORDS(lngID)
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    
EXITPROCEDURE:
    Exit Sub
    
errHandler:
    MsgBox Err.Description, vbCritical, "Agency Delete"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub

Private Sub cmdEditBank_Click()
    blnAdd = False
    blnEdit = True
    Call fn_UNSET_CONTROL_COLOR("Bank")
    Call subEnableCtl
    txtShortName.SetFocus
End Sub



Private Sub cmdSaveBank_Click()
On Error GoTo errHandler

    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtBankName, "Please enter bank name.") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtBranch, "Please enter bank branch.") Then Exit Sub
    
    With cls_BANK_Obj
        .ShortName = txtShortName.Text
        .BankName = txtBankName.Text
        .Branch = txtBranch.Text
        
        If blnAdd = True And blnEdit = False Then
            Call .fn_SAVE_BANK_RECORDS
            Else
                Call .fn_UPDATE_BANK_RECORDS(lngID)
        End If
    End With
    
    Call sub_EMPTY_FIELS("Bank")
    Call fn_SET_CONTROL_COLOR("Bank")
    Call subDisableCtl
    Call subLoadBank
    
EXITPROCEDURE:
    Exit Sub
    
errHandler:
    MsgBox Err.Description, vbCritical, "SaveBank"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub

Private Sub cmdCancelAccount_Click()
    Call disableAccountCtl
    fra.Enabled = False
'    fraFeatures.Enabled = False
    Call loadBankAccounts(lngID)
    Call sub_EMPTY_FIELS("Account")
    Call fn_SET_CONTROL_COLOR("Account")
End Sub

Private Sub cmdCloseAccount_Click()
    Unload Me
End Sub

Private Sub cmdDeleteAccount_Click()
'On Error GoTo errHandler
'
'    If MsgBox("Are you sure you want to delete  " & Trim(txtUserName.Text) & " ?", vbQuestion + vbYesNo, Title) = vbNo Then
'        Exit Sub
'        Else
'            Call cls_USER_Obj.fn_DELETE_USER(lstUsers.ItemData(lstUsers.ListIndex))
'            MsgBox "Users successfully deleted.", vbInformation, Title
'    End If
'    sub_EMPTY_FIELS ("Account")
'    Call loadUsers(lngID)
'
'EXITPROCEDURE:
'    Exit Sub
'
'errHandler:
'    MsgBox Err.Description, vbCritical, "Connection"
'    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
'    GoTo EXITPROCEDURE
End Sub

Private Sub cmdEditAccount_Click()
    Call fn_UNSET_CONTROL_COLOR("Account")
    fra.Enabled = True
    blnEditAccount = True
    blnAddAccount = False
    Call enableAccountCtl
    fra.Enabled = True
    txtAccountNo.SetFocus
End Sub



Private Sub cmdNewAccount_Click()
    
    If Trim(txtBankName.Text) = "" Then
        MsgBox "Please select the bank name", vbExclamation, "New Account"
        Exit Sub
    End If
    
    Call fn_UNSET_CONTROL_COLOR("Account")
    blnEditAccount = False
    blnAddAccount = True
    Call enableAccountCtl
    fra.Enabled = True
    cmdEditAccount.Enabled = False
    cmdDeleteAccount.Enabled = False
    Call sub_EMPTY_FIELS("Account")

    txtAccountNo.SetFocus

End Sub

Private Sub cmdSaveAccount_Click()
On Error GoTo errHandler

    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtAccountNo, "Please enter account number.") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtAccountName, "Please enter account name.") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_COMBO(cboAccountType, "Please select account type.") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtOpeningBalance, "Please enter openeing balance.") Then Exit Sub
    
    Dim ctr As Long
    If Mdl_FUNCTIONS.fn_FILL_ALL_FIEL(Me) = True Then
        MsgBox "All the fields are required", vbInformation, "Save"
        Exit Sub
    End If
    
    
    With cls_BANK_Obj
        .BankID = lngID
        .AccountNo = Trim(txtAccountNo.Text)
        .AccountName = Trim(txtAccountName.Text)
        .AccountTypeID = cboAccountType.ItemData(cboAccountType.ListIndex)
        .OpenDate = DTPOpenDate.Value
        .OpenBalance = Trim(txtOpeningBalance.Text)
    End With
    
    
    If blnAddAccount = True Then
            Call cls_BANK_Obj.fn_SAVE_BANK_ACCOUNT_RECORDS
        Else
            Call cls_BANK_Obj.fn_UPDATE_BANK_ACCOUNT_RECORDS(lstBankAccount.ItemData(lstBankAccount.ListIndex))
    End If
    
    
    Call fn_SET_CONTROL_COLOR("Account")
    Call disableAccountCtl

    Call sub_EMPTY_FIELS("Account")

    Call loadBankAccounts(lngID)
    
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
    Call subLoadBank
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    Call subDisableCtl
    Call subLoadAccountTypes
    Call disableAccountCtl
    
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdNewBank, cmdNewBank.hwnd, "Add new bank details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdEditBank, cmdEditBank.hwnd, "Edit existing bank details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdDeleteBank, cmdDeleteBank.hwnd, "Delete existing bank details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSaveBank, cmdSaveBank.hwnd, "Save bank details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancelBank, cmdCancelBank.hwnd, "Cancel the transaction.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
'    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdPrint, cmdPrint.hwnd, "Print the list of all the employee", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCloseBank, cmdCloseBank.hwnd, "Close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdNewAccount, cmdNewAccount.hwnd, "Add new account.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdEditAccount, cmdEditAccount.hwnd, "Edit existing account details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdDeleteAccount, cmdDeleteAccount.hwnd, "Delete existing account details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSaveAccount, cmdSaveAccount.hwnd, "Save account details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancelAccount, cmdCancelAccount.hwnd, "Cancel the transaction.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
'    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdPrint, cmdPrint.hwnd, "Print the list of all the employee", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCloseAccount, cmdCloseAccount.hwnd, "Close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    
End Sub

Public Sub subLoadAccountTypes()

    Dim rec As New ADODB.Recordset
    Set rec = cls_ACCOUNT_TYPE_Obj.fn_LOAD_ACCOUNT_TYPE(0)
    
    cboAccountType.Clear
    
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            cboAccountType.AddItem rec!AccountType
            cboAccountType.ItemData(cboAccountType.NewIndex) = rec!AccountTypeID
            rec.MoveNext
        Loop
    End If
    
End Sub

Private Sub subLoadBank()

    Dim rec As New ADODB.Recordset
    Set rec = cls_BANK_Obj.fn_LOAD_BANKS(0)
    
    lstBank.Clear
    
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            lstBank.AddItem rec!BankName
            lstBank.ItemData(lstBank.NewIndex) = rec!BankID
            rec.MoveNext
        Loop
    End If
    
End Sub

Private Sub subLoadBankDetails(lngID As Long)

    Dim rec As New ADODB.Recordset
    Set rec = cls_BANK_Obj.fn_LOAD_BANKS(lngID)
    
    If rec.AbsolutePosition <> -1 Then
        txtShortName.Text = Trim(rec!ShortName)
        txtBankName.Text = Trim(rec!BankName)
        txtBranch.Text = Trim(rec!Branch)
        cmdEditBank.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blnAdd = False
    blnEdit = False
    blnAccountType = False
End Sub


Private Sub lstBank_Click()
    lngID = lstBank.ItemData(lstBank.ListIndex)
    Call sub_EMPTY_FIELS("Bank")
    Call subLoadBankDetails(lngID)
    Call loadBankAccounts(lngID)
End Sub

Private Sub subDisableCtl()

    cmdNewBank.Enabled = True
    cmdCloseBank.Enabled = True
    
    cmdEditBank.Enabled = False
    cmdSaveBank.Enabled = False
    cmdDeleteBank.Enabled = False
    cmdCancelBank.Enabled = False
    
    fraBankDetails.Enabled = False
    lstBank.Enabled = True
End Sub

Private Sub subEnableCtl()

    cmdNewBank.Enabled = False
    cmdCloseBank.Enabled = False
    
    cmdEditBank.Enabled = False
    cmdSaveBank.Enabled = True
    cmdDeleteBank.Enabled = False
    cmdCancelBank.Enabled = True
    
    fraBankDetails.Enabled = True
    lstBank.Enabled = False
End Sub


Private Sub loadBankAccounts(lngID As Long)

    Dim rec As New ADODB.Recordset
    
    Set rec = cls_BANK_Obj.fn_LOAD_BANKS_ACCOUNTS(lngID)
    
    lstBankAccount.Clear
    
    If rec.AbsolutePosition = -1 Then Exit Sub
    Do While Not rec.EOF
        lstBankAccount.AddItem rec!AccountName
        lstBankAccount.ItemData(lstBankAccount.NewIndex) = rec!AccountID
        rec.MoveNext
    Loop

End Sub

Private Sub loadAccountDetails(lngUserID As Long)

    Dim rec As New ADODB.Recordset
    
    Set rec = cls_BANK_Obj.fn_LOAD_BANKS_ACCOUNTS_RECORDS(lngAccountID)
    
    If rec.AbsolutePosition = -1 Then Exit Sub
    txtAccountNo.Text = Trim(rec!AccountNo)
    txtAccountName.Text = Trim(rec!AccountName)
    cboAccountType.ListIndex = Mdl_FUNCTIONS.fn_GET_LIST_INDEX(cboAccountType, rec!AccountTypeID)
    txtOpeningBalance.Text = Trim(rec!OpenBalance)
    DTPOpenDate.Value = Trim(rec!OpenDate)

    cmdEditAccount.Enabled = True
'    cmdDeleteUser.Enabled = True
    
End Sub

Private Sub lstbankaccount_Click()

    lngAccountID = lstBankAccount.ItemData(lstBankAccount.ListIndex)
    Call loadAccountDetails(lstBankAccount.ItemData(lstBankAccount.ListIndex))
    
End Sub


Private Sub disableAccountCtl()

    cmdNewAccount.Enabled = True
    cmdSaveAccount.Enabled = False
    cmdDeleteAccount.Enabled = False
    cmdCancelAccount.Enabled = False
    cmdEditAccount.Enabled = False

End Sub

Private Sub enableAccountCtl()

    cmdNewAccount.Enabled = False
    cmdSaveAccount.Enabled = True
    cmdDeleteAccount.Enabled = True
    cmdCancelAccount.Enabled = True
    cmdEditAccount.Enabled = True

End Sub


Public Sub sub_EMPTY_FIELS(str As String)

    If str = "Bank" Then
        txtShortName.Text = ""
        txtBankName.Text = ""
        txtBranch.Text = ""

        Else
            txtAccountNo.Text = ""
            txtAccountName.Text = ""
            cboAccountType.ListIndex = -1
            txtOpeningBalance.Text = ""
            DTPOpenDate.Value = Date
    End If

End Sub

Public Function fn_SET_CONTROL_COLOR(str As String)
    Dim col As Variant
    col = &H8000000F

    If str = "Bank" Then
        txtShortName.BackColor = col
        txtBankName.BackColor = col
        txtBranch.BackColor = col

        Else
            txtAccountNo.BackColor = col
            txtAccountName.BackColor = col
            cboAccountType.BackColor = col
            txtOpeningBalance.BackColor = col
            
    End If
    

End Function

Public Function fn_UNSET_CONTROL_COLOR(str As String)
    Dim col As Variant
    col = &HFFFFFF

    If str = "Bank" Then
        txtShortName.BackColor = col
        txtBankName.BackColor = col
        txtBranch.BackColor = col

        Else
            txtAccountNo.BackColor = col
            txtAccountName.BackColor = col
            cboAccountType.BackColor = col
            txtOpeningBalance.BackColor = col
            
    End If
    
End Function

Private Sub txtAgencyName_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_ALPHABET_ONLY(KeyAscii)
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtFullName_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_ALPHABET_ONLY(KeyAscii)
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
'    Call Mdl_FUNCTIONS.sub_ALPHABET_ONLY(KeyAscii)
End Sub

Private Sub txtOpeningBalance_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub
