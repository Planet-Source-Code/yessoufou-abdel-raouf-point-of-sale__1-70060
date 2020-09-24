VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_PRODUCTS 
   Caption         =   "PRODUCTS"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   12180
   Begin OCX.b8Container b8Container1 
      Height          =   9075
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   16007
      BackColor       =   16185592
      Begin OCX.b8Container b8Container5 
         Height          =   3405
         Left            =   90
         TabIndex        =   21
         Top             =   4440
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   6006
         BackColor       =   16185592
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Price Control"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   3225
            Left            =   180
            TabIndex        =   22
            Top             =   60
            Width           =   11595
            Begin VB.Frame fraPrice 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   525
               Left            =   5970
               TabIndex        =   31
               Top             =   2640
               Width           =   5505
               Begin lvButton.lvButtons_H cmdManagePackages 
                  Height          =   495
                  Left            =   60
                  TabIndex        =   32
                  Top             =   0
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   873
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
                  Image           =   "frm_PRODUCTS.frx":0000
                  ImgSize         =   24
                  cBack           =   -2147483633
                  mPointer        =   99
                  mIcon           =   "frm_PRODUCTS.frx":0515
               End
               Begin lvButton.lvButtons_H cmdEditPriceControl 
                  Height          =   495
                  Left            =   2085
                  TabIndex        =   33
                  Top             =   0
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   873
                  Caption         =   "&Edit ..."
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
                  Image           =   "frm_PRODUCTS.frx":082F
                  ImgSize         =   24
                  cBack           =   -2147483633
                  mPointer        =   99
                  mIcon           =   "frm_PRODUCTS.frx":0BAC
               End
               Begin lvButton.lvButtons_H cmdDeletePrice 
                  Height          =   495
                  Left            =   4035
                  TabIndex        =   34
                  Top             =   0
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
                  Image           =   "frm_PRODUCTS.frx":0EC6
                  ImgSize         =   24
                  cBack           =   -2147483633
                  mPointer        =   99
                  mIcon           =   "frm_PRODUCTS.frx":13D9
               End
            End
            Begin MSComctlLib.ListView lvw 
               Height          =   2295
               Left            =   120
               TabIndex        =   23
               Top             =   300
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   4048
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
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
               NumItems        =   8
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "PackageID"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Package"
                  Object.Width           =   2823
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   2
                  Text            =   "Qty"
                  Object.Width           =   1765
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   3
                  Text            =   "Supplier Price"
                  Object.Width           =   3529
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   4
                  Text            =   "Selling Price(-Tax)"
                  Object.Width           =   3529
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   5
                  Text            =   "VAT"
                  Object.Width           =   1766
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   6
                  Text            =   "NHIL"
                  Object.Width           =   1766
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   7
                  Text            =   "Selling Price(+Tax)"
                  Object.Width           =   4763
               EndProperty
            End
         End
      End
      Begin OCX.b8Container b8Container4 
         Height          =   4365
         Left            =   3120
         TabIndex        =   8
         Top             =   60
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7699
         BackColor       =   16185592
         Begin VB.Frame fraProductDetails 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Height          =   4215
            Left            =   90
            TabIndex        =   9
            Top             =   90
            Width           =   8745
            Begin VB.Frame Frame1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Product Details"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   4185
               Left            =   30
               TabIndex        =   12
               Top             =   -30
               Width           =   8655
               Begin VB.TextBox txtInitialStock 
                  Alignment       =   2  'Center
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
                  Left            =   1500
                  MaxLength       =   100
                  TabIndex        =   4
                  Top             =   1980
                  Width           =   2415
               End
               Begin VB.TextBox txtReorderLevel 
                  Alignment       =   2  'Center
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
                  Left            =   1500
                  MaxLength       =   100
                  TabIndex        =   6
                  Top             =   2820
                  Width           =   2415
               End
               Begin VB.TextBox txtUnitsInStock 
                  Alignment       =   2  'Center
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
                  Left            =   1500
                  MaxLength       =   100
                  TabIndex        =   5
                  Top             =   2400
                  Width           =   2415
               End
               Begin VB.CheckBox ChkActive 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Left            =   1500
                  TabIndex        =   3
                  Top             =   1620
                  Width           =   255
               End
               Begin VB.ComboBox cboCategories 
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
                  Left            =   1500
                  Style           =   2  'Dropdown List
                  TabIndex        =   1
                  Top             =   780
                  Width           =   3375
               End
               Begin VB.TextBox txtProductName 
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
                  Left            =   1500
                  MaxLength       =   100
                  TabIndex        =   0
                  Top             =   360
                  Width           =   5025
               End
               Begin VB.ComboBox cboPackages 
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
                  Left            =   1500
                  Style           =   2  'Dropdown List
                  TabIndex        =   2
                  Top             =   1200
                  Width           =   3375
               End
               Begin lvButton.lvButtons_H cboLoadCategories 
                  Height          =   375
                  Left            =   4920
                  TabIndex        =   17
                  Top             =   750
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
                  Image           =   "frm_PRODUCTS.frx":16F3
                  ImgSize         =   24
                  cBack           =   -2147483633
                  mPointer        =   99
                  mIcon           =   "frm_PRODUCTS.frx":1C08
               End
               Begin lvButton.lvButtons_H cmdLoadPackages 
                  Height          =   375
                  Left            =   4920
                  TabIndex        =   18
                  Top             =   1170
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
                  Image           =   "frm_PRODUCTS.frx":1F22
                  ImgSize         =   24
                  cBack           =   -2147483633
                  mPointer        =   99
                  mIcon           =   "frm_PRODUCTS.frx":2437
               End
               Begin VB.Label Label6 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Initial Stock"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   150
                  TabIndex        =   36
                  Top             =   1980
                  Width           =   1305
               End
               Begin VB.Label Label5 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Reorder Level"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   150
                  TabIndex        =   20
                  Top             =   2820
                  Width           =   1305
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Units In Stock"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   150
                  TabIndex        =   19
                  Top             =   2400
                  Width           =   1305
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Active"
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
                  Left            =   150
                  TabIndex        =   16
                  Top             =   1650
                  Width           =   795
               End
               Begin VB.Label Label3 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Main Package"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   150
                  TabIndex        =   15
                  Top             =   1230
                  Width           =   1335
               End
               Begin VB.Label Label4 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Categories"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   120
                  TabIndex        =   14
                  Top             =   780
                  Width           =   1125
               End
               Begin VB.Label Label2 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Product Name"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   150
                  TabIndex        =   13
                  Top             =   360
                  Width           =   1305
               End
            End
         End
      End
      Begin OCX.b8Container b8Container2 
         Height          =   4365
         Left            =   90
         TabIndex        =   10
         Top             =   60
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   7699
         BackColor       =   16185592
         Begin VB.ComboBox cboProducts 
            Appearance      =   0  'Flat
            Height          =   3690
            Left            =   120
            Style           =   1  'Simple Combo
            TabIndex        =   35
            Top             =   510
            Width           =   2805
         End
         Begin OCX.b8SideTab b8SideTab1 
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   150
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   661
            Caption         =   "Products"
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
         Left            =   90
         TabIndex        =   24
         Top             =   7860
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1191
         BackColor       =   16185592
         Begin lvButton.lvButtons_H cmdAddNewProduct 
            Height          =   495
            Left            =   180
            TabIndex        =   25
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
            Image           =   "frm_PRODUCTS.frx":2751
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_PRODUCTS.frx":2C66
         End
         Begin lvButton.lvButtons_H cmdCancelProduct 
            Height          =   495
            Left            =   8244
            TabIndex        =   26
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
            Image           =   "frm_PRODUCTS.frx":2F80
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_PRODUCTS.frx":34EF
         End
         Begin lvButton.lvButtons_H cmdEditProduct 
            Height          =   495
            Left            =   2196
            TabIndex        =   27
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
            Image           =   "frm_PRODUCTS.frx":3809
            ImgSize         =   24
            Enabled         =   0   'False
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_PRODUCTS.frx":3B86
         End
         Begin lvButton.lvButtons_H cmdCloseProduct 
            Height          =   495
            Left            =   10260
            TabIndex        =   28
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
            Image           =   "frm_PRODUCTS.frx":3EA0
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_PRODUCTS.frx":4430
         End
         Begin lvButton.lvButtons_H cmdSaveProduct 
            Height          =   495
            Left            =   6228
            TabIndex        =   29
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
            Image           =   "frm_PRODUCTS.frx":474A
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_PRODUCTS.frx":49AA
         End
         Begin lvButton.lvButtons_H cmdDeleteProduct 
            Height          =   495
            Left            =   4212
            TabIndex        =   30
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
            Image           =   "frm_PRODUCTS.frx":4CC4
            ImgSize         =   24
            Enabled         =   0   'False
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_PRODUCTS.frx":51D7
         End
      End
   End
End
Attribute VB_Name = "frm_PRODUCTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngProductID As Long
Dim strPackageName As String

Private Sub cboLoadCategories_Click()
    blnCategory = True
    frm_CATEGORIES.Show
End Sub

Private Sub cmdAddNewProduct_Click()
    Call Mdl_FUNCTIONS.fn_UNSET_CONTROL_COLOR(Me, "Product")
    
    blnAddProduct = True
    blnEditProduct = False
    
    Call sub_EMPTY_FIELDS("Product")
    Call fn_ENABLE_CONTROLS("Product")
    txtProductName.SetFocus
End Sub

Private Sub cmdCancelProduct_Click()
    Call sub_EMPTY_FIELDS("Product")
    Call fn_DISABLE_CONTROLS("Product")
    Call sub_LOAD_PRODUCTS
    Call fn_SET_CONTROL_COLOR(Me, "Product")
End Sub

Private Sub cmdCloseProduct_Click()
    Unload Me
End Sub

Private Sub cmdDeletePrice_Click()
    If lvw.ListItems.Count = 0 Then Exit Sub
    
    If MsgBox("Do you Realy want to delete thise details?", 4 + 32, Title) = vbYes Then
        lvw.ListItems.Remove lvw.SelectedItem.Index
    End If
End Sub

Private Sub cmdEditPriceControl_Click()
    
    If lvw.ListItems.Count = 0 Then Exit Sub
    
    blnEditPriceControl = True
    With frm_ADD_PACKAGES
        .cboPackages.ListIndex = Mdl_FUNCTIONS.fn_GET_LIST_INDEX(cboPackages, lvw.SelectedItem.Text)
        .txtQty.Text = lvw.SelectedItem.ListSubItems(2).Text
        .txtSupplierPrice.Text = lvw.SelectedItem.ListSubItems(3).Text
        .txtSellingPriceWithoutTax.Text = lvw.SelectedItem.ListSubItems(4).Text
        .txtVat.Text = lvw.SelectedItem.ListSubItems(5).Text
        If Val(.txtVat.Text) > 0 Then .ChkVAT.Value = 0
        .txtNHIL.Text = lvw.SelectedItem.ListSubItems(6).Text
        If Val(.txtNHIL.Text) > 0 Then .ChkNHIL.Value = 0
        .txtSellingPriceWithTax.Text = lvw.SelectedItem.ListSubItems(7).Text
        lvw.ListItems.Remove lvw.SelectedItem.Index
        .Show 1
    End With
    
    
    
End Sub

Private Sub cmdEditProduct_Click()

    If txtProductName.Text = "" Then
        MsgBox "Please select the product to Edit."
        Exit Sub
    End If
        
    Call fn_UNSET_CONTROL_COLOR(Me, "Product")
    
    blnAddProduct = False
    blnEditProduct = True
    
    Call fn_ENABLE_CONTROLS("Product")
    txtProductName.SetFocus
End Sub

Private Sub cmdLoadPackages_Click()
    blnProductPackage = True
    frm_PACKAGES.Show 1
End Sub

Private Sub cmdManagePackages_Click()
    frm_ADD_PACKAGES.Show 1
End Sub

Private Sub cmdSaveProduct_Click()
On Error GoTo errHandler

    Dim ctr As Long
    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtProductName, "Please enter the product name!") Then Exit Sub
    If Mdl_FUNCTIONS.fn_REQUIRE_COMBO_FIELD(cboCategories, "Please select the category of the product!") Then Exit Sub

    With cls_PRODUCT_Obj
        .ProductID = .fn_AUTOGEN
        .ProductName = Trim(txtProductName.Text)
        .CategoryID = cboCategories.ItemData(cboCategories.ListIndex)
        .PackageID = cboPackages.ItemData(cboPackages.ListIndex)
        .InitialStock = Val(Trim(txtInitialStock.Text))
        .UnitsInStock = Val(Trim(txtUnitsInStock.Text))
        .ReOrderLevel = Val(Trim(txtReOrderLevel.Text))
        
        If ChkActive.Value = 0 Then
            .Active = 1
            Else
                .Active = 0
        End If

        If blnAddProduct = True And blnEditProduct = False Then
            Call .fn_SAVE_PRODUCTS_RECORDS
            Else
                Call .fn_UPDATE_PRODUCTS_RECORDS(lngProductID)
        End If
    End With
    
    Call cls_PRODUCT_PACKAGE_Obj.fn_DELETE_PRODUCT_PACKAGE(lngProductID)
    
    For ctr = 1 To lvw.ListItems.Count
        With cls_PRODUCT_PACKAGE_Obj
            .ProductID = lngProductID
            .PackageID = lvw.ListItems(ctr).Text
            .Qty = lvw.ListItems(ctr).ListSubItems(2).Text
            .SupplierPrice = lvw.ListItems(ctr).ListSubItems(3).Text
            .SellingPriceWithoutTax = lvw.ListItems(ctr).ListSubItems(4).Text
            .VAT = lvw.ListItems(ctr).ListSubItems(5).Text
            .NHIL = lvw.ListItems(ctr).ListSubItems(6).Text
            .SellingPriceWithTax = lvw.ListItems(ctr).ListSubItems(7).Text
            .fn_SAVE_PRODUCTS_PACKAGE
        End With
    Next
     
    Call sub_EMPTY_FIELDS("Product")
    Call sub_LOAD_PRODUCTS
    Call fn_DISABLE_CONTROLS("Product")
    Call fn_SET_CONTROL_COLOR(Me, "Product")
    
EXITPROCEDURE:
    Exit Sub

errHandler:
    MsgBox Err.Description, vbCritical, Title
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_FORM_SIZE(Me)
    Move 0, 0
    
    Call fn_SET_CONTROL_COLOR(Me, "All")
    Call fn_DISABLE_CONTROLS("Category")
    Call fn_DISABLE_CONTROLS("Product")
    
    Call sub_LOAD_CATEGORIES(cboCategories)
    Call sub_LOAD_PACKAGES
    Call sub_LOAD_PRODUCTS
    
    
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdAddNewProduct, cmdAddNewProduct.hwnd, "Add new product details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdEditProduct, cmdEditProduct.hwnd, "Edit existing product details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdDeleteProduct, cmdDeleteProduct.hwnd, "Delete existing product details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSaveProduct, cmdSaveProduct.hwnd, "Save product details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancelProduct, cmdCancelProduct.hwnd, "Cancel the transaction.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCloseProduct, cmdCloseProduct.hwnd, "Close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
   
End Sub

Public Sub sub_LOAD_CATEGORIES(cbo As ComboBox)
    Dim rec As New ADODB.Recordset
    
    Set rec = cls_CATEGORY_Obj.fN_LOAD_CATEGORIES(0)

    cbo.Clear
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            cbo.AddItem rec!CategoryName
            cbo.ItemData(cbo.NewIndex) = rec!CategoryID
            rec.MoveNext
        Loop
    End If
    
End Sub

Public Sub sub_LOAD_PACKAGES()
    Dim rec As New ADODB.Recordset
    
    Set rec = cls_REFERENCES_Obj.fn_LOAD_PACKAGES(0)
    
    cboPackages.Clear
    
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            cboPackages.AddItem rec!PackageName
            cboPackages.ItemData(cboPackages.NewIndex) = rec!PackageID
            rec.MoveNext
        Loop
    End If
    
End Sub

Private Sub sub_LOAD_PRODUCTS()
    Dim rec As New ADODB.Recordset
    
    Set rec = cls_PRODUCT_Obj.fN_LOAD_PRODUCTS_DETAILS()
    cboProducts.Clear
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            cboProducts.AddItem rec!ProductName
            cboProducts.ItemData(cboProducts.NewIndex) = rec!ProductID
            rec.MoveNext
        Loop
    End If
    
End Sub

Private Sub sub_LOAD_PRODUCTS_DETAILS(lngProductID As Long)
    Dim rec As New ADODB.Recordset
    
    Set rec = cls_PRODUCT_Obj.fN_LOAD_PRODUCTS_DETAILS(lngProductID)

    If rec.AbsolutePosition <> -1 Then
        txtProductName.Text = Trim(rec!ProductName) & ""
        cboCategories.ListIndex = Mdl_FUNCTIONS.fn_GET_LIST_INDEX(cboCategories, rec!CategoryID)
        cboPackages.ListIndex = Mdl_FUNCTIONS.fn_GET_LIST_INDEX(cboPackages, rec!PackageID)
        txtInitialStock.Text = Trim(rec!InitialStock)
        txtUnitsInStock.Text = Trim(rec!UnitsInStock)
        txtReOrderLevel.Text = Trim(rec!ReOrderLevel)
        If CLng(rec!Active) = 0 Then
            ChkActive.Value = 1
            Else
                ChkActive.Value = 0
        End If
        
        cmdEditProduct.Enabled = True
        cmdDeleteProduct.Enabled = True
        
    End If
    
End Sub

Private Sub sub_LOAD_PRODUCT_PACKAGE(lngProductID As Long)
    Dim rec As New ADODB.Recordset
    
    Set rec = cls_PRODUCT_PACKAGE_Obj.fn_LOAD_PRODUCT_PACKAGE(lngProductID)
    
    lvw.ListItems.Clear
    
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            Set lstItem = lvw.ListItems.Add(, , rec!PackageID)
                Call sub_LOAD_PACKAGES_DETAILS(rec!PackageID)
                lstItem.ListSubItems.Add , , strPackageName
                lstItem.ListSubItems.Add , , rec!Qty
                lstItem.ListSubItems.Add , , rec!SupplierPrice
                lstItem.ListSubItems.Add , , rec!SellingPriceWithoutTax
                lstItem.ListSubItems.Add , , rec!VAT
                lstItem.ListSubItems.Add , , rec!NHIL
                lstItem.ListSubItems.Add , , rec!SellingPriceWithTax
            rec.MoveNext
        Loop
    End If
    
End Sub

Private Sub sub_LOAD_PACKAGES_DETAILS(lngPackageID As Long)

    Dim rec As New ADODB.Recordset
    Set rec = cls_REFERENCES_Obj.fn_LOAD_PACKAGES(lngPackageID)
    
    If rec.AbsolutePosition <> -1 Then
        strPackageName = rec!PackageName
    End If
End Sub

Private Sub cboProducts_Click()
    If cboProducts.ListIndex = -1 Then Exit Sub
    
    lngProductID = cboProducts.ItemData(cboProducts.ListIndex)
    Call sub_LOAD_PRODUCTS_DETAILS(lngProductID)
    Call sub_LOAD_PRODUCT_PACKAGE(lngProductID)
End Sub

'**********Function to disable some buttons***********
Private Function fn_DISABLE_CONTROLS(str As String)

    If str = "Category" Then

        
    ElseIf str = "Product" Then
    
        cmdSaveProduct.Enabled = False
        cmdCancelProduct.Enabled = False
            
        cmdEditProduct.Enabled = True
        cmdDeleteProduct.Enabled = True
        cmdAddNewProduct.Enabled = True
        fraProductDetails.Enabled = False
        cboProducts.Enabled = True
        fraProductDetails.Enabled = False
        fraPrice.Enabled = False
    End If
    
End Function

'**********Function to enable some buttons***********
Private Function fn_ENABLE_CONTROLS(str As String)

    If str = "Category" Then

        
    ElseIf str = "Product" Then
    
        cmdSaveProduct.Enabled = True
        cmdCancelProduct.Enabled = True
            
        cmdEditProduct.Enabled = False
        cmdDeleteProduct.Enabled = False
        cmdAddNewProduct.Enabled = False
        fraProductDetails.Enabled = True
        cboProducts.Enabled = False
        fraProductDetails.Enabled = True
        fraPrice.Enabled = True
    
    End If

End Function

Private Sub sub_EMPTY_FIELDS(str As String)

    If str = "Category" Then
        
        
    ElseIf str = "Product" Then
        txtUnitsInStock.Text = ""
        txtReOrderLevel.Text = ""
        txtProductName.Text = ""
        txtInitialStock.Text = ""
        cboCategories.ListIndex = -1
        cboPackages.ListIndex = -1
        ChkActive.Value = 0
        lvw.ListItems.Clear
    End If

End Sub


Private Sub txtInitialStock_Change()
'    txtUnitsInStock.Text = Val(txtInitialStock.Text)
End Sub

Private Sub txtReorderLevel_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtUnitsInStock_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub
