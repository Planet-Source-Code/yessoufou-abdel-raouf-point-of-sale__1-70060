VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_PENDING_SALES 
   Caption         =   "PENDING SALES"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   12180
   Begin OCX.b8Container b8Container1 
      Height          =   8565
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   15108
      BackColor       =   16185592
      Begin OCX.b8Container b8Container3 
         Height          =   4125
         Left            =   90
         TabIndex        =   1
         Top             =   3690
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   7276
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
            Top             =   210
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
            Left            =   9150
            TabIndex        =   18
            Top             =   3630
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
            Left            =   7920
            TabIndex        =   17
            Top             =   3660
            Width           =   1335
         End
      End
      Begin OCX.b8Container b8Container2 
         Height          =   3615
         Left            =   90
         TabIndex        =   4
         Top             =   90
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   6376
         BackColor       =   16185592
         Begin VB.ComboBox cbo 
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
            TabIndex        =   5
            Text            =   "cboProductName"
            Top             =   480
            Width           =   3765
         End
         Begin OCX.b8SideTab b8SideTab1 
            Height          =   375
            Left            =   150
            TabIndex        =   6
            Top             =   150
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   661
            Caption         =   "Sales"
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
         Left            =   4200
         TabIndex        =   7
         Top             =   90
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   6376
         BackColor       =   16185592
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
            TabIndex        =   8
            Top             =   630
            Width           =   5565
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
            TabIndex        =   13
            Top             =   645
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
            TabIndex        =   12
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
            TabIndex        =   11
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
            TabIndex        =   10
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
            TabIndex        =   9
            Top             =   210
            Width           =   1245
         End
      End
      Begin OCX.b8Container b8Container5 
         Height          =   675
         Left            =   90
         TabIndex        =   14
         Top             =   7830
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   1191
         BackColor       =   16185592
         Begin lvButton.lvButtons_H cmdSave 
            Height          =   495
            Left            =   8070
            TabIndex        =   15
            Top             =   90
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
            Image           =   "frm_PENDING_SALES.frx":0000
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_PENDING_SALES.frx":0260
         End
         Begin lvButton.lvButtons_H cmdCancel 
            Height          =   495
            Left            =   9900
            TabIndex        =   16
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
            Image           =   "frm_PENDING_SALES.frx":057A
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_PENDING_SALES.frx":0B0A
         End
      End
   End
End
Attribute VB_Name = "frm_PENDING_SALES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dblGrossAmount As Double
Public dblDiscount As Double
Public dblNetAmount As Double
Public dblAmountPaid As Double
Public dblBalance As Double

Dim lngSalesID As Long
Private Sub sub_LOAD_SALES(lngID As Long)
    Dim rec As New ADODB.Recordset
    
    Set rec = cls_SALES_Obj.fn_LOAD_PENDING_SALES(lngID)
    cbo.Clear
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            cbo.AddItem rec!SalesNo
            cbo.ItemData(cbo.NewIndex) = rec!SalesID
            rec.MoveNext
        Loop
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim lngOrderID As Long
    Dim ctr As Long
    
    With frm_PENDING_SALES_PAYMENT
        blnPendingSales = True
        .LblGrossAmount.Caption = lblTotalAmount.Caption
        .Show 1
    End With
    
    If blnPendingSales = False Then Exit Sub
    Call cls_SALES_Obj.fn_UPDATE_SALES_STATUS(lngSalesID)
    
    With cls_RECEIPTS_Obj
        .SalesID = lngSalesID
        .GrossAmount = dblGrossAmount
        .Discount = dblDiscount
        .NetAmount = dblNetAmount
        .AmountPaid = dblAmountPaid
        .Balance = dblBalance
        .fn_SAVE_RECEIPT_RECORDS
        With frm_SALES_PAYMENT_REP
            .lngSalesID = lngSalesID
            .Show
        End With
    End With
        
'    MsgBox "Transaction saved successfully", vbInformation, Title
    Call sub_CLEAR_FIELD
    lvw.ListItems.Clear
    lblTotalAmount.Caption = ""
    Call sub_LOAD_SALES(0)

End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_FORM_SIZE(Me)
    Move 0, 0
    Call sub_LOAD_SALES(0)
    Call sub_LOAD_CUSTOMERS
    
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSave, cmdSave.hwnd, "Tender credit sales details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancel, cmdCancel.hwnd, "Cancel and close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor


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
Private Sub cbo_Click()
    If cbo.ListIndex = -1 Then Exit Sub
    
    lngSalesID = cbo.ItemData(cbo.ListIndex)
    Call sub_LOAD_SALES_DETAILS(lngSalesID)
    Call sub_LOAD_PRODUCTS_DETAILS(lngSalesID)
End Sub

Private Sub sub_LOAD_SALES_DETAILS(lngSalesID As Long)
    Dim rec As New ADODB.Recordset
    
    Set rec = cls_SALES_Obj.fn_LOAD_PENDING_SALES(lngSalesID)
    If rec.AbsolutePosition <> -1 Then
        lblSalesNo.Caption = Trim(rec!SalesNo) & ""
        lblSalesDate.Caption = Trim(rec!SalesDate) & ""
        If Trim(rec!CustomerID) = 0 Then
            cboCustomer.ListIndex = -1
            Else
                cboCustomer.ListIndex = Mdl_FUNCTIONS.fn_GET_LIST_INDEX(cboCustomer, rec!CustomerID)
        End If
        lblTotalAmount.Caption = rec!Total
    End If
    
End Sub

Private Sub sub_LOAD_PRODUCTS_DETAILS(lngSalesID As Long)
    Dim rec As New ADODB.Recordset
    
    lvw.ListItems.Clear
    
    Set rec = cls_SALES_Obj.fn_LOAD_PENDING_SALES(lngSalesID)
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            Set lstItem = lvw.ListItems.Add(, , rec!ProductID)
                lstItem.ListSubItems.Add , , rec!ProductName
                lstItem.ListSubItems.Add , , rec!PackageName
                lstItem.ListSubItems.Add , , rec!SellingPrice
                lstItem.ListSubItems.Add , , rec!Qty
                lstItem.ListSubItems.Add , , rec!Expr1
            rec.MoveNext
        Loop
        
    End If
    
End Sub

Private Sub sub_CLEAR_FIELD()
    
    lblSalesDate.Caption = ""
    lblSalesNo.Caption = ""
    cboCustomer.ListIndex = -1

End Sub
