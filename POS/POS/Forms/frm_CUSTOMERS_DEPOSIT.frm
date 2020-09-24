VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_CUSTOMERS_DEPOSIT 
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
      Height          =   9075
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   16007
      BackColor       =   16185592
      Begin OCX.b8Container b8Container2 
         Height          =   4635
         Left            =   120
         TabIndex        =   1
         Top             =   150
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   8176
         BackColor       =   16185592
         Begin VB.Frame fraDetails 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   4395
            Left            =   2940
            TabIndex        =   3
            Top             =   90
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
               TabIndex        =   12
               Top             =   690
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
               Height          =   765
               Left            =   1695
               MaxLength       =   200
               MultiLine       =   -1  'True
               TabIndex        =   11
               Text            =   "frm_CUSTOMERS_DEPOSIT.frx":0000
               Top             =   3480
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
               TabIndex        =   10
               Text            =   " "
               Top             =   3090
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
               TabIndex        =   9
               Top             =   2685
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
               ItemData        =   "frm_CUSTOMERS_DEPOSIT.frx":0002
               Left            =   1695
               List            =   "frm_CUSTOMERS_DEPOSIT.frx":000C
               TabIndex        =   8
               Top             =   2280
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
               TabIndex        =   7
               Text            =   " "
               Top             =   1875
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
               TabIndex        =   6
               Text            =   " "
               Top             =   1470
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
               TabIndex        =   5
               Text            =   " "
               Top             =   1080
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
               TabIndex        =   4
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
               TabIndex        =   21
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
               TabIndex        =   20
               Top             =   1080
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
               TabIndex        =   19
               Top             =   1890
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
               TabIndex        =   18
               Top             =   1470
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
               TabIndex        =   17
               Top             =   705
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
               TabIndex        =   16
               Top             =   2280
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
               TabIndex        =   15
               Top             =   2685
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
               TabIndex        =   14
               Top             =   3480
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
               TabIndex        =   13
               Top             =   3090
               Width           =   1455
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
            Height          =   3930
            Left            =   150
            TabIndex        =   2
            Top             =   540
            Width           =   2715
         End
         Begin OCX.b8Container fra 
            Height          =   4305
            Left            =   9180
            TabIndex        =   22
            Top             =   180
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   7594
            BackColor       =   16185592
            Begin VB.PictureBox Picture1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   2685
               Left            =   150
               ScaleHeight     =   2655
               ScaleWidth      =   2265
               TabIndex        =   23
               Top             =   210
               Width           =   2295
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
                  TabIndex        =   24
                  Top             =   1110
                  Width           =   1965
               End
               Begin VB.Image imgPicture 
                  Height          =   2655
                  Left            =   0
                  Picture         =   "frm_CUSTOMERS_DEPOSIT.frx":001E
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   2265
               End
            End
            Begin lvButton.lvButtons_H cmdIdentification 
               Height          =   375
               Left            =   150
               TabIndex        =   25
               Top             =   2880
               Width           =   2310
               _ExtentX        =   4075
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
         End
         Begin OCX.b8SideTab b8SideTab1 
            Height          =   375
            Left            =   150
            TabIndex        =   26
            Top             =   180
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
   End
End
Attribute VB_Name = "frm_CUSTOMERS_DEPOSIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

