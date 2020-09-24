VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_ORDERS 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ORDERS"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   12180
   Begin OCX.b8Container b8Container4 
      Height          =   3225
      Left            =   0
      TabIndex        =   0
      Top             =   5220
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   5689
      BackColor       =   16185592
      Begin VB.TextBox txtTotalQty 
         Height          =   465
         Left            =   3240
         TabIndex        =   1
         Top             =   3870
         Width           =   2355
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   2595
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   11565
         _ExtentX        =   20399
         _ExtentY        =   4577
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
            Text            =   "Buying Unit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Supplier Price"
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
         Left            =   9120
         TabIndex        =   4
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label8 
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
         Left            =   7890
         TabIndex        =   3
         Top             =   2790
         Width           =   1335
      End
   End
   Begin OCX.b8Container b8Container2 
      Height          =   5235
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   9234
      BackColor       =   16185592
      Begin VB.ListBox lstSuppliers 
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
         Height          =   3930
         Left            =   120
         TabIndex        =   29
         Top             =   510
         Width           =   2715
      End
      Begin OCX.b8Container b8Container3 
         Height          =   4275
         Left            =   2910
         TabIndex        =   6
         Top             =   180
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   7541
         BackColor       =   16185592
         Begin VB.ComboBox cboProductName 
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
            Height          =   3885
            Left            =   120
            Style           =   1  'Simple Combo
            TabIndex        =   18
            Text            =   "cboProductName"
            Top             =   180
            Width           =   2505
         End
         Begin VB.Frame fraProductDetails 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1185
            Left            =   2610
            TabIndex        =   11
            Top             =   1020
            Width           =   5685
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
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   390
               Width           =   4125
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
               Left            =   1440
               MaxLength       =   100
               TabIndex        =   13
               Top             =   0
               Width           =   4095
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
               Left            =   1440
               MaxLength       =   11
               TabIndex        =   12
               Top             =   780
               Width           =   2115
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
               TabIndex        =   17
               Top             =   390
               Width           =   1005
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
               TabIndex        =   16
               Top             =   -30
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
               TabIndex        =   15
               Top             =   780
               Width           =   1245
            End
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
            Left            =   4050
            MaxLength       =   11
            TabIndex        =   10
            Top             =   2550
            Width           =   2145
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
            Left            =   4050
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2190
            Width           =   4125
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
            Left            =   4050
            MaxLength       =   11
            TabIndex        =   8
            Top             =   2970
            Width           =   2145
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
            Left            =   4050
            MaxLength       =   7
            TabIndex        =   7
            Top             =   3360
            Width           =   2145
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Order No"
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
            Left            =   2670
            TabIndex        =   28
            Top             =   270
            Width           =   1245
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Order Date"
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
            Left            =   2670
            TabIndex        =   27
            Top             =   630
            Width           =   1365
         End
         Begin VB.Label lblOrderNo 
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
            Left            =   4050
            TabIndex        =   26
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label lblOrderDate 
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
            Left            =   4050
            TabIndex        =   25
            Top             =   630
            Width           =   4095
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   " Qty/No"
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
            Left            =   2640
            TabIndex        =   24
            Top             =   2550
            Width           =   1125
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Cost"
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
            Left            =   2700
            TabIndex        =   23
            Top             =   2970
            Width           =   1395
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Buying Unit"
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
            Left            =   2700
            TabIndex        =   22
            Top             =   2190
            Width           =   1515
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
            Left            =   2700
            TabIndex        =   21
            Top             =   3360
            Width           =   1395
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
            Left            =   4050
            TabIndex        =   20
            Top             =   3780
            Width           =   2145
         End
         Begin VB.Label Label10 
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
            Left            =   2700
            TabIndex        =   19
            Top             =   3780
            Width           =   1335
         End
      End
      Begin OCX.b8Container b8Container5 
         Height          =   675
         Left            =   90
         TabIndex        =   30
         Top             =   4470
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   1191
         BackColor       =   16185592
         Begin lvButton.lvButtons_H cmdRemove 
            Height          =   495
            Left            =   2475
            TabIndex        =   31
            Top             =   90
            Width           =   2025
            _ExtentX        =   3572
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
            Image           =   "frm_ORDERS.frx":0000
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_ORDERS.frx":0513
         End
         Begin lvButton.lvButtons_H cmdProcess 
            Height          =   495
            Left            =   210
            TabIndex        =   32
            Top             =   90
            Width           =   2025
            _ExtentX        =   3572
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
            Image           =   "frm_ORDERS.frx":082D
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_ORDERS.frx":0D42
         End
         Begin lvButton.lvButtons_H cmdClearList 
            Height          =   495
            Left            =   4755
            TabIndex        =   33
            Top             =   90
            Width           =   2025
            _ExtentX        =   3572
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
            Image           =   "frm_ORDERS.frx":105C
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_ORDERS.frx":146C
         End
         Begin lvButton.lvButtons_H cmdSave 
            Height          =   495
            Left            =   7050
            TabIndex        =   34
            Top             =   90
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   873
            Caption         =   "&Save Order"
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
            Image           =   "frm_ORDERS.frx":1786
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_ORDERS.frx":19E6
         End
         Begin lvButton.lvButtons_H cmdCancel 
            Height          =   495
            Left            =   9300
            TabIndex        =   35
            Top             =   90
            Width           =   2025
            _ExtentX        =   3572
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
            Image           =   "frm_ORDERS.frx":1D00
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_ORDERS.frx":2290
         End
      End
      Begin OCX.b8SideTab b8SideTab2 
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   180
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   661
         Caption         =   "Suppliers"
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
Attribute VB_Name = "frm_ORDERS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngSupplierID As Long
Dim lngProductID As Long
Dim lngItemPosition As Long
Dim strPackageName As String
Dim lngPackageID As Long

Private Sub cboProductName_Click()
    If cboProductName.ListIndex = -1 Then Exit Sub
    
    lngProductID = cboProductName.ItemData(cboProductName.ListIndex)
    Call sub_LOAD_PRODUCTS_DETAILS(lngProductID)
    Call sub_LOAD_PRODUCT_PACKAGE(lngProductID)
End Sub

Private Sub sub_LOAD_PRODUCTS_DETAILS(lngProductID As Long)
    Dim rec As New ADODB.Recordset
    
    Set rec = cls_PRODUCT_Obj.fN_LOAD_PRODUCTS_DETAILS(lngProductID)

    If rec.AbsolutePosition <> -1 Then
        txtProductName.Text = Trim(rec!ProductName) & ""
        cboCategories.ListIndex = Mdl_FUNCTIONS.fn_GET_LIST_INDEX(cboCategories, rec!CategoryID)
'        dblStockAvailable = Trim(Rec!UnitsInStock)
        txtUnitInStock.Text = Trim(rec!UnitsInStock) & ""
'        txtReOrderLevel.Text = Trim(Rec!ReOrderLevel) & ""
    End If
    
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

Private Sub cboSellingUnit_Click()
    If cboSellingUnit.ListIndex = -1 Then Exit Sub
    lngPackageID = cboSellingUnit.ItemData(cboSellingUnit.ListIndex)
    Call sub_LOAD_PRODUCT_PACKAGE_DETAILS(lngProductID, lngPackageID)
End Sub

Private Sub sub_LOAD_PRODUCT_PACKAGE_DETAILS(lngProductID As Long, lngPackageID As Long)

    Dim rec As New ADODB.Recordset
    Set rec = cls_PRODUCT_PACKAGE_Obj.fn_LOAD_PRODUCT_PACKAGE_DETAILS(lngProductID, lngPackageID)
    
    If rec.AbsolutePosition <> -1 Then
        txtSellingPrice.Text = Trim(rec!SupplierPrice)
        txtQty.Text = Trim(rec!Qty)
    End If
    
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

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClearList_Click()
    If lvw.ListItems.Count > 0 Then
        If MsgBox("Are you sure you want to clear the list?", 4 + 32, Title) = vbYes Then
            lvw.ListItems.Clear
            lblTotalAmount.Caption = ""
            lstSuppliers.Enabled = True
        End If
    End If
End Sub

Private Sub cmdProcess_Click()

    If cboProductName.ListIndex = -1 Then Exit Sub
    
    If Mdl_FUNCTIONS.fn_REQUIRE_COMBO_FIELD(cboSellingUnit, "Please select the selling Unit") = True Then Exit Sub
    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtQtySold, "Please enter the quantity ordered") = True Then Exit Sub
    
'    If Val(txtTotalQty.Text) > Val(Trim(txtUnitInStock.Text)) Then
'        MsgBox "Units in stock is not enough", vbInformation, Title
'        txtQtySold.SetFocus
'        Exit Sub
'    End If
    
    If fn_CHECK_IF_PRODUCT_EXIST(lngProductID, lngPackageID) = True Then
            lvw.ListItems(lngItemPosition).ListSubItems(4).Text = Val(lvw.ListItems(lngItemPosition).ListSubItems(4).Text) + Val(txtQtySold.Text)
            lvw.ListItems(lngItemPosition).ListSubItems(5).Text = Val(lvw.ListItems(lngItemPosition).ListSubItems(5).Text) + Val(lblTotal.Caption)
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
    
    lstSuppliers.Enabled = False
    
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
    
    lblTotalAmount.Caption = Val(lblTotalAmount.Caption) - Val(lvw.SelectedItem.ListSubItems(5).Text)
    lvw.ListItems.Remove lvw.SelectedItem.Index
    
    If lvw.ListItems.Count = 0 Then
        lstSuppliers.Enabled = True
    End If
    
End Sub

Private Sub cmdSave_Click()
    Dim lngOrderID As Long
    Dim ctr As Long
    
    If lvw.ListItems.Count = 0 Then Exit Sub
    If MsgBox("Are you sure you want to save transaction?", vbQuestion + vbYesNo, Title) = vbYes Then
        With cls_ORDER_Obj
            lngOrderID = .fn_ID_AUTOGEN
            .OrderID = lngOrderID
            .OrderNo = .fn_AUTOGEN
            .SupplierID = lngSupplierID
            .OrderDate = Date
            .OrderTime = Now
            .TotalOrder = Val(lblTotalAmount.Caption)
            .Status = 0
            .fn_SAVE_ORDERS_RECORDS
        End With
        
        For ctr = 1 To lvw.ListItems.Count
            With cls_ORDER_Obj
                .OrderID = lngOrderID
                .ProductID = lvw.ListItems(ctr).Text
                .BuyingUnit = lvw.ListItems(ctr).ListSubItems(6).Text
                .SupplierPrice = lvw.ListItems(ctr).ListSubItems(3).Text
                .Qty = lvw.ListItems(ctr).ListSubItems(4).Text
                .Total = lvw.ListItems(ctr).ListSubItems(5).Text
                .fn_SAVE_ORDERS_DETAILS_RECORDS
            End With
        Next
            
        MsgBox "Transaction saved successfully", vbInformation, Title
        Call sub_CLEAR_FIELD
        lvw.ListItems.Clear
        lblTotalAmount.Caption = ""
        lstSuppliers.Enabled = True
    End If
    
    
End Sub



Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_FORM_SIZE(Me)
    Move 0, 0
    lblOrderNo.Caption = cls_ORDER_Obj.fn_AUTOGEN
    lblOrderDate.Caption = Now
    
    Call sub_LOAD_SUPPLIERS(lstSuppliers)
    Call sub_LOAD_CATEGORIES(cboCategories)
    
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdProcess, cmdProcess.hwnd, "Add order details to the list.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
'    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdEdit, cmdEdit.hwnd, "Edit existing order details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdRemove, cmdRemove.hwnd, "Remove order details from the list.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSave, cmdSave.hwnd, "Save order details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancel, cmdCancel.hwnd, "Cancel and close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdClearList, cmdClearList.hwnd, "Clear the order list.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor

    
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

Private Sub lstSuppliers_Click()

    If lstSuppliers.ListIndex = -1 Then Exit Sub

    txtUnitInStock.Text = ""
    txtSellingPrice.Text = ""
    lvw.ListItems.Clear
    lblTotalAmount.Caption = ""
    lngSupplierID = lstSuppliers.ItemData(lstSuppliers.ListIndex)
    Call sub_LOAD_SUPPLIER_PRODUCT(lngSupplierID)
    
End Sub

Private Sub txtQtysold_Change()
    If cboProductName.ListIndex = -1 Then Exit Sub
    lblTotal.Caption = Val(txtQtySold.Text) * Val(txtSellingPrice.Text)
    txtTotalQty.Text = Val(txtQty.Text) * Val(txtQtySold.Text)
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub sub_LOAD_SUPPLIERS(Optional lst As ListBox)

    Dim rec As New ADODB.Recordset
    Set rec = cls_SUPPLIER_Obj.fn_LOAD_SUPPLIERS(0)
    
    lst.Clear
    
    Do While Not rec.EOF
        lst.AddItem rec!CompanyName
        lst.ItemData(lst.NewIndex) = rec!SupplierID
        rec.MoveNext
    Loop

End Sub

Private Sub sub_LOAD_SUPPLIER_PRODUCT(lngSupplierID As Long)

    Dim rec As New ADODB.Recordset
    Set rec = cls_PRODUCT_Obj.fn_LOAD_ACTIVE_PRODUCT(lngSupplierID)
    
    cboProductName.Clear
    
    Do While Not rec.EOF
        cboProductName.AddItem rec!ProductName
        cboProductName.ItemData(cboProductName.NewIndex) = rec!ProductID
        rec.MoveNext
    Loop

End Sub

Private Sub sub_CLEAR_FIELD()
    
    cboProductName.ListIndex = -1
    txtUnitInStock.Text = ""
    txtSellingPrice.Text = ""
    txtQty.Text = ""
    lblTotal.Caption = ""
    txtQtySold.Text = ""
    
End Sub
