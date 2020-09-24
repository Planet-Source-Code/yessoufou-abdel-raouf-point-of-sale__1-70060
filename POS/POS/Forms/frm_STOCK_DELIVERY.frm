VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_STOCK_DELIVERY 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin OCX.b8Container b8Container1 
      Height          =   8385
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   14790
      BackColor       =   16185592
      Begin VB.Frame fraDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2505
         Left            =   2970
         TabIndex        =   3
         Top             =   60
         Width           =   9045
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
            TabIndex        =   8
            Text            =   " "
            Top             =   300
            Width           =   3885
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
            TabIndex        =   7
            Text            =   " "
            Top             =   1080
            Width           =   3885
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
            TabIndex        =   6
            Text            =   " "
            Top             =   1470
            Width           =   3855
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
            TabIndex        =   5
            Text            =   " "
            Top             =   1875
            Width           =   3855
         End
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
            TabIndex        =   4
            Top             =   690
            Width           =   1935
         End
         Begin OCX.b8Container fra 
            Height          =   2265
            Left            =   6960
            TabIndex        =   21
            Top             =   150
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   3995
            BackColor       =   16185592
            Begin VB.PictureBox Picture1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   1845
               Left            =   150
               ScaleHeight     =   1815
               ScaleWidth      =   1635
               TabIndex        =   22
               Top             =   210
               Width           =   1665
               Begin VB.Image imgPicture 
                  Height          =   1815
                  Left            =   0
                  Picture         =   "frm_STOCK_DELIVERY.frx":0000
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   1635
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
            TabIndex        =   13
            Top             =   705
            Width           =   1305
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
            TabIndex        =   12
            Top             =   1470
            Width           =   1455
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
            TabIndex        =   11
            Top             =   1875
            Width           =   1485
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
            TabIndex        =   10
            Top             =   1080
            Width           =   1215
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
            TabIndex        =   9
            Top             =   300
            Width           =   1305
         End
      End
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
         TabIndex        =   1
         Top             =   510
         Width           =   2715
      End
      Begin OCX.b8SideTab b8SideTab1 
         Height          =   375
         Left            =   150
         TabIndex        =   2
         Top             =   150
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   661
         Caption         =   "Customers"
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
      TabIndex        =   14
      Top             =   8400
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   1191
      BackColor       =   16185592
      Begin lvButton.lvButtons_H cmdAddNew 
         Height          =   495
         Left            =   180
         TabIndex        =   15
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
         Image           =   "frm_STOCK_DELIVERY.frx":26B8
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_STOCK_DELIVERY.frx":2BCD
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   8436
         TabIndex        =   16
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
         Image           =   "frm_STOCK_DELIVERY.frx":2EE7
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_STOCK_DELIVERY.frx":3456
      End
      Begin lvButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   2244
         TabIndex        =   17
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
         Image           =   "frm_STOCK_DELIVERY.frx":3770
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_STOCK_DELIVERY.frx":3AED
      End
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   495
         Left            =   10500
         TabIndex        =   18
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
         Image           =   "frm_STOCK_DELIVERY.frx":3E07
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_STOCK_DELIVERY.frx":4397
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   6372
         TabIndex        =   19
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
         Image           =   "frm_STOCK_DELIVERY.frx":46B1
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_STOCK_DELIVERY.frx":4911
      End
      Begin lvButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   4308
         TabIndex        =   20
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
         Image           =   "frm_STOCK_DELIVERY.frx":4C2B
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_STOCK_DELIVERY.frx":513E
      End
   End
End
Attribute VB_Name = "frm_STOCK_DELIVERY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

