VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_ORDERS_FROM_CUSTOMERS 
   BorderStyle     =   0  'None
   Caption         =   "Orders From Customers"
   ClientHeight    =   9090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   Begin OCX.b8Container b8Container1 
      Height          =   9075
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   16007
      BackColor       =   16185592
      Begin OCX.b8Container b8Container3 
         Height          =   4275
         Left            =   150
         TabIndex        =   1
         Top             =   4620
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   7541
         BackColor       =   16185592
         Begin VB.TextBox txtTotalQty 
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
            Height          =   465
            Left            =   4950
            TabIndex        =   2
            Top             =   4440
            Width           =   2235
         End
         Begin MSComctlLib.ListView lvw 
            Height          =   3375
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   11565
            _ExtentX        =   20399
            _ExtentY        =   5953
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
               Text            =   "ProductID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ProductName"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Selling Unit"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Selling Price"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Quantity"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Total"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "PackageID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "TotalQty"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label lblTotalAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   9180
            TabIndex        =   5
            Top             =   3630
            Width           =   2505
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount"
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
            Left            =   7920
            TabIndex        =   4
            Top             =   3630
            Width           =   1335
         End
      End
      Begin OCX.b8Container b8Container2 
         Height          =   3615
         Left            =   150
         TabIndex        =   6
         Top             =   180
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   6376
         BackColor       =   16185592
         Begin VB.ComboBox cboProductName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2910
            Left            =   150
            Style           =   1  'Simple Combo
            TabIndex        =   7
            Text            =   "cboProductName"
            Top             =   480
            Width           =   3765
         End
         Begin OCX.b8SideTab b8SideTab1 
            Height          =   375
            Left            =   150
            TabIndex        =   8
            Top             =   150
            Width           =   3765
            _ExtentX        =   6641
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
      Begin OCX.b8Container b8Container4 
         Height          =   3615
         Left            =   4260
         TabIndex        =   9
         Top             =   180
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   6376
         BackColor       =   16185592
         Begin VB.Frame fraProductDetails 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1275
            Left            =   90
            TabIndex        =   15
            Top             =   510
            Width           =   7455
            Begin VB.TextBox txtReOrderLevel 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   4890
               MaxLength       =   11
               TabIndex        =   19
               Top             =   930
               Width           =   2115
            End
            Begin VB.TextBox txtUnitInStock 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   1470
               MaxLength       =   11
               TabIndex        =   18
               Top             =   930
               Width           =   2025
            End
            Begin VB.TextBox txtProductName 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   1470
               MaxLength       =   100
               TabIndex        =   17
               Top             =   90
               Width           =   5535
            End
            Begin VB.ComboBox cboCategories 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   1470
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   510
               Width           =   5565
            End
            Begin VB.Label Label9 
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
               Height          =   315
               Left            =   3630
               TabIndex        =   23
               Top             =   930
               Width           =   1305
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Unit In Stock"
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
               Left            =   90
               TabIndex        =   22
               Top             =   930
               Width           =   1245
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
               Index           =   0
               Left            =   90
               TabIndex        =   21
               Top             =   90
               Width           =   1305
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
               Left            =   90
               TabIndex        =   20
               Top             =   510
               Width           =   1005
            End
         End
         Begin VB.TextBox txtQtySold 
            Alignment       =   1  'Right Justify
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
            Height          =   345
            Left            =   1560
            MaxLength       =   7
            TabIndex        =   14
            Top             =   3090
            Width           =   2175
         End
         Begin VB.ComboBox cboCustomer 
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1860
            Width           =   5565
         End
         Begin VB.TextBox txtSellingPrice 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1560
            MaxLength       =   11
            TabIndex        =   12
            Top             =   2670
            Width           =   2175
         End
         Begin VB.ComboBox cboSellingUnit 
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
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2280
            Width           =   5565
         End
         Begin VB.TextBox txtQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   4950
            MaxLength       =   11
            TabIndex        =   10
            Top             =   2670
            Width           =   2145
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount"
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
            Left            =   3780
            TabIndex        =   34
            Top             =   3150
            Width           =   1335
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   345
            Left            =   4950
            TabIndex        =   33
            Top             =   3090
            Width           =   2145
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity"
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
            Index           =   2
            Left            =   180
            TabIndex        =   32
            Top             =   3090
            Width           =   1395
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name"
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
            TabIndex        =   31
            Top             =   1875
            Width           =   1545
         End
         Begin VB.Label lblSalesDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   4860
            TabIndex        =   30
            Top             =   180
            Width           =   2235
         End
         Begin VB.Label lblSalesNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   1560
            TabIndex        =   29
            Top             =   180
            Width           =   2115
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Date"
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
            Left            =   3810
            TabIndex        =   28
            Top             =   180
            Width           =   1365
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Sales No"
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
            Left            =   180
            TabIndex        =   27
            Top             =   210
            Width           =   1245
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Selling Unit"
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
            Left            =   180
            TabIndex        =   26
            Top             =   2280
            Width           =   1515
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Selling Price"
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
            Left            =   180
            TabIndex        =   25
            Top             =   2670
            Width           =   1395
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Selling Qty"
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
            Left            =   3780
            TabIndex        =   24
            Top             =   2700
            Width           =   1125
         End
      End
      Begin OCX.b8Container b8Container5 
         Height          =   675
         Left            =   150
         TabIndex        =   35
         Top             =   3870
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   1191
         BackColor       =   16185592
         Begin lvButton.lvButtons_H cmdRemove 
            Height          =   495
            Left            =   2632
            TabIndex        =   36
            Top             =   90
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            Caption         =   "&Remove From List"
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
            Image           =   "frm_ORDERS_FROM_CUSTOMERS.frx":0000
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_ORDERS_FROM_CUSTOMERS.frx":0513
         End
         Begin lvButton.lvButtons_H cmdProcess 
            Height          =   495
            Left            =   210
            TabIndex        =   37
            Top             =   90
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            Caption         =   "&Add To List"
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
            Image           =   "frm_ORDERS_FROM_CUSTOMERS.frx":082D
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_ORDERS_FROM_CUSTOMERS.frx":0D42
         End
         Begin lvButton.lvButtons_H cmdClearList 
            Height          =   495
            Left            =   5054
            TabIndex        =   38
            Top             =   90
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            Caption         =   "&Clear List"
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
            Image           =   "frm_ORDERS_FROM_CUSTOMERS.frx":105C
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_ORDERS_FROM_CUSTOMERS.frx":146C
         End
         Begin lvButton.lvButtons_H cmdSave 
            Height          =   495
            Left            =   6030
            TabIndex        =   39
            Top             =   690
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            Caption         =   "&Tender"
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
            Image           =   "frm_ORDERS_FROM_CUSTOMERS.frx":1786
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_ORDERS_FROM_CUSTOMERS.frx":19E6
         End
         Begin lvButton.lvButtons_H cmdCancel 
            Height          =   495
            Left            =   9900
            TabIndex        =   40
            Top             =   90
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            Caption         =   "&Close"
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
            Image           =   "frm_ORDERS_FROM_CUSTOMERS.frx":1D00
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_ORDERS_FROM_CUSTOMERS.frx":2290
         End
         Begin lvButton.lvButtons_H cmdHold 
            Height          =   495
            Left            =   7476
            TabIndex        =   41
            Top             =   90
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            Caption         =   "&Print Receipt"
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
            Image           =   "frm_ORDERS_FROM_CUSTOMERS.frx":25AA
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_ORDERS_FROM_CUSTOMERS.frx":28C4
         End
      End
   End
End
Attribute VB_Name = "frm_ORDERS_FROM_CUSTOMERS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngProductID As Long
Dim lngSupplierID As Long
Dim lngOrderID As Long
Dim lngCustomerID As Long
Dim lngPackageID As Long
Dim lngItemPosition As Long
Dim strPackageName As String

Dim dblQtySold As Double
Dim dblStockAvailable As Double

Public dblGrossAmount As Double
Public dblDiscount As Double
Public dblNetAmount As Double
Public dblAmountPaid As Double
Public dblBalance As Double

Dim strCompanyName As String

Dim blnCheckStock As Boolean

Private Sub sub_LOAD_PRODUCTS(lngCategoryID As Long)
    Dim rec As New ADODB.Recordset
    
    Set rec = cls_PRODUCT_Obj.fN_LOAD_ACTIVE_PRODUCTS(lngCategoryID)
    cboProductName.Clear
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            cboProductName.AddItem rec!ProductName
            cboProductName.ItemData(cboProductName.NewIndex) = rec!ProductID
            rec.MoveNext
        Loop
    End If
    
End Sub

Private Sub cboCustomer_Click()
    If cboCustomer.ListIndex = -1 Then
        lngCustomerID = 0
        Else
            lngCustomerID = cboCustomer.ItemData(cboCustomer.ListIndex)
    End If
End Sub

Private Sub cboProductName_Click()
    If cboProductName.ListIndex = -1 Then Exit Sub
    
    lngProductID = cboProductName.ItemData(cboProductName.ListIndex)
    Call sub_CLEAR_FIELD
    Call sub_LOAD_PRODUCTS_DETAILS(lngProductID)
    Call sub_LOAD_PRODUCT_PACKAGE(lngProductID)
    
    
End Sub

Private Sub sub_LOAD_PRODUCT_PACKAGE(lngProductID As Long)
    Dim rec As New ADODB.Recordset
    
    Set rec = cls_PRODUCT_PACKAGE_Obj.fn_LOAD_PRODUCT_PACKAGE(lngProductID)

    cboSellingUnit.Clear

    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            Call sub_LOAD_PACKAGES_DETAILS(rec!PackageID)
            cboSellingUnit.AddItem strPackageName
            cboSellingUnit.ItemData(cboSellingUnit.NewIndex) = rec!PackageID
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


Private Sub sub_LOAD_PRODUCT_PACKAGE_DETAILS(lngPackageID As Long)

    Dim rec As New ADODB.Recordset
    Set rec = cls_PRODUCT_PACKAGE_Obj.fn_LOAD_PRODUCT_PACKAGE_DETAILS(lngProductID, lngPackageID)
    
    If rec.AbsolutePosition <> -1 Then
        txtSellingPrice.Text = Trim(rec!SellingPriceWithTax)
        txtQty.Text = Trim(rec!Qty)
    End If
End Sub

Private Sub cboSellingUnit_Click()
    If cboSellingUnit.ListIndex = -1 Then Exit Sub
    lngPackageID = cboSellingUnit.ItemData(cboSellingUnit.ListIndex)
    Call sub_LOAD_PRODUCT_PACKAGE_DETAILS(lngPackageID)
    txtQtySold.Text = ""
End Sub

Private Sub cmdCancel_Click()
    If lvw.ListItems.Count = 0 Then
        Unload Me
        Exit Sub
    End If
    If MsgBox("Are you sure you want to cancel transactions ?", 4 + 32, Title) = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdClearList_Click()
    If lvw.ListItems.Count > 0 Then
        If MsgBox("Are you sure you want to clear the list?", 4 + 32, Title) = vbYes Then
            lvw.ListItems.Clear
        End If
    End If
End Sub

Private Sub cmdHold_Click()
    blnPending = True
    cmdSave_Click
End Sub

Private Sub cmdProcess_Click()

    If cboProductName.ListIndex = -1 Then Exit Sub
    
    If Mdl_FUNCTIONS.fn_REQUIRE_COMBO_FIELD(cboSellingUnit, "Please select the selling Unit") = True Then Exit Sub
    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtQtySold, "Please enter the quantity ordered") = True Then Exit Sub
    
    
    If Val(txtTotalQty.Text) > Val(Trim(txtUnitInStock.Text)) Then
        MsgBox "Units in stock is not enough", vbInformation, Title
        txtQtySold.SetFocus
        Exit Sub
    End If
    
    If fn_CHECK_IF_PRODUCT_EXIST(lngProductID, lngPackageID) = True Then
            lvw.ListItems(lngItemPosition).ListSubItems(4).Text = Val(lvw.ListItems(lngItemPosition).ListSubItems(4).Text) + Val(txtQtySold.Text)
            lvw.ListItems(lngItemPosition).ListSubItems(5).Text = Val(lvw.ListItems(lngItemPosition).ListSubItems(5).Text) + Val(lblTotal.Caption)
            lvw.ListItems(lngItemPosition).ListSubItems(7).Text = Val(lvw.ListItems(lngItemPosition).ListSubItems(7).Text) + Val(txtTotalQty.Text)
        Else
            Set lstItem = lvw.ListItems.Add(, , lngProductID)
              lstItem.ListSubItems.Add , , cboProductName.Text
              lstItem.ListSubItems.Add , , Trim(cboSellingUnit.Text)
              lstItem.ListSubItems.Add , , Trim(txtSellingPrice.Text)
              lstItem.ListSubItems.Add , , Trim(txtQtySold.Text)
              lstItem.ListSubItems.Add , , Trim(lblTotal.Caption)
              lstItem.ListSubItems.Add , , Trim(lngPackageID)
              lstItem.ListSubItems.Add , , Trim(txtTotalQty.Text)
    End If
    
    lblTotalAmount.Caption = Val(lblTotalAmount.Caption) + Val(lblTotal.Caption)
    
    Call sub_CLEAR_FIELD

End Sub




Private Function fn_CHECK_IF_PRODUCT_EXIST(lngProductID As Long, lngSellingUnit As Long) As Boolean

    Dim ctr As Long
    fn_CHECK_IF_PRODUCT_EXIST = False
    For ctr = 1 To lvw.ListItems.Count
        If lngProductID = lvw.ListItems(ctr).Text And lngSellingUnit = lvw.ListItems(ctr).ListSubItems(6).Text Then
            fn_CHECK_IF_PRODUCT_EXIST = True
            lngItemPosition = ctr
        End If
    Next

End Function


Private Sub cmdRemove_Click()
    If lvw.ListItems.Count = 0 Then Exit Sub
    
    lblTotalAmount.Caption = Val(lblTotalAmount.Caption) - Val(lvw.SelectedItem.ListSubItems(4).Text)
    lvw.ListItems.Remove lvw.SelectedItem.Index

End Sub

Private Sub cmdSave_Click()
    Dim lngOrderID As Long
    Dim ctr As Long
    
    If Mdl_FUNCTIONS.fn_REQUIRE_COMBO_FIELD(cboCustomer, "Please select the customer who is ordering") = True Then Exit Sub
    
    If lvw.ListItems.Count = 0 Then
        MsgBox "Please Add Sales Details", vbExclamation, Title
        Exit Sub
    End If
    
    If blnPending = False Then
        With frm_SALES_PAYMENT
            .LblGrossAmount.Caption = lblTotalAmount.Caption
            .Show 1
        End With
    End If
        
    With cls_CUSTOMERS_ORDERS_Obj
        lngOrderID = .fn_ID_AUTOGEN
        .OrderID = lngOrderID
        .OrderNo = .fn_AUTOGEN
        .CustomerID = lngCustomerID
        .OrderDate = Date
        .OrderTime = Now
        .OrdersTotal = Val(lblTotalAmount.Caption)
        If blnPending = True Then
            .Status = 1
            Else
                .Status = 0
        End If
    End With

    For ctr = 1 To lvw.ListItems.Count
        With cls_CUSTOMERS_ORDERS_Obj
            .OrderID = lngOrderID
            .ProductID = lvw.ListItems(ctr).Text
            Call fn_GET_TOTAL(.ProductID, lvw.ListItems(ctr).ListSubItems(1).Text)
            If blnCheckStock = True Then Exit Sub
            .SellingUnit = lvw.ListItems(ctr).ListSubItems(6).Text
            .SellingPrice = lvw.ListItems(ctr).ListSubItems(3).Text
            .Qty = lvw.ListItems(ctr).ListSubItems(4).Text
            .Total = lvw.ListItems(ctr).ListSubItems(5).Text
            .UnitsSold = .Qty
            .fn_SAVE_ORDERS_DETAILS_RECORDS
            .fn_UPDATE_PRODUCTS_IN_STOCK (.ProductID)
        End With
    Next
    
    cls_CUSTOMERS_ORDERS_Obj.fn_SAVE_ORDERS_RECORDS
    
'    If blnPending = False Then
'        With cls_RECEIPTS_Obj
'            .SalesID = lngSalesID
'            .GrossAmount = dblGrossAmount
'            .Discount = dblDiscount
'            .NetAmount = dblNetAmount
'            .AmountPaid = dblAmountPaid
'            .Balance = dblBalance
'            .fn_SAVE_RECEIPT_RECORDS
'            With frm_SALES_PAYMENT_REP
'                .lngSalesID = lngSalesID
'                .Show
'            End With
'        End With
'     End If
        
'    MsgBox "Transaction saved successfully", vbInformation, Title
    Call sub_CLEAR_FIELD
    lvw.ListItems.Clear
    lblTotalAmount.Caption = ""
    blnPending = False
        
End Sub

Private Function fn_GET_TOTAL(lngProductID As Long, strProductName As String)

    Dim ctr As Long
    blnCheckStock = False
    dblQtySold = 0
    For ctr = 1 To lvw.ListItems.Count
        If lngProductID = lvw.ListItems(ctr).Text Then
            dblQtySold = dblQtySold + Val(lvw.ListItems(ctr).ListSubItems(7).Text)
        End If
    Next
    
    If fn_GET_TOTAL_STOCK(lngProductID) < dblQtySold Then
        MsgBox "Product " & strProductName & " Is Not Enough", vbInformation, Title
        blnCheckStock = True
    End If
    
End Function

Private Function fn_GET_TOTAL_STOCK(lngProductID As Long) As Double

    Dim rec As New ADODB.Recordset
    
    Set rec = cls_PRODUCT_Obj.fN_LOAD_PRODUCTS_DETAILS(lngProductID)

    fn_GET_TOTAL_STOCK = 0

    If rec.AbsolutePosition <> -1 Then
        fn_GET_TOTAL_STOCK = rec!UnitsInStock
    End If
    
End Function
'Private Function fn_CHECK_STOCK(lngProductID As Long, dblStock As Double) As Boolean
'
'    Dim ctr As Long
'    fn_CHECK_STOCK = False
'    For ctr = 1 To lvw.ListItems.Count
'        If lngProductID = lvw.ListItems(ctr).Text Then
'            dblQtySold = dblQtySold + Val(lvw.ListItems(ctr))
'            If dblQtySold > dblStock Then
'                fn_CHECK_STOCK = True
'            End If
'        End If
'    Next
'
'
'
'End Function

Private Sub Form_Load()
    Move 0, 0
    Call sub_LOAD_PRODUCTS(0)
    Call sub_LOAD_CATEGORIES(cboCategories)
    Call sub_LOAD_CUSTOMERS
'    Call sub_LOAD_COMPANY_INFO
    lblSalesNo.Caption = cls_SALES_Obj.fn_AUTOGEN
    lblSalesDate.Caption = Now
    blnPending = False
End Sub

Private Sub sub_LOAD_CUSTOMERS()

    Dim rec As New ADODB.Recordset
    Set rec = cls_CUSTOMER_Obj.fn_LOAD_CUSTOMERS(0)

    cboCustomer.Clear

    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            cboCustomer.AddItem rec!FirstName & " " & rec!LastName
            cboCustomer.ItemData(cboCustomer.NewIndex) = rec!CustomerID
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
        dblStockAvailable = Trim(rec!UnitsInStock)
        txtUnitInStock.Text = Trim(rec!UnitsInStock) & ""
        txtReOrderLevel.Text = Trim(rec!ReOrderLevel) & ""
    End If
    
End Sub

Private Sub sub_LOAD_CATEGORIES(cbo As ComboBox)
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

Private Sub sub_CLEAR_FIELD()
    
    txtProductName.Text = ""
    cboCategories.ListIndex = -1
    cboSellingUnit.ListIndex = -1
    txtSellingPrice.Text = ""
'    txtUnitOnOrder.Text = ""
    txtUnitInStock.Text = ""
    txtReOrderLevel.Text = ""
    txtQty.Text = ""
    txtQtySold.Text = ""
    lblTotal.Caption = ""
    txtTotalQty.Text = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)
    blnPending = False
End Sub

Private Sub txtQtysold_Change()
    If cboProductName.ListIndex = -1 Then Exit Sub
    lblTotal.Caption = Val(txtQtySold.Text) * Val(txtSellingPrice.Text)
    txtTotalQty.Text = Val(txtQty.Text) * Val(txtQtySold.Text)
'    txtUnitInStock.Text = Val(Val(txtSellingPrice.Text) - Val(txtQtySold.Text))
End Sub



