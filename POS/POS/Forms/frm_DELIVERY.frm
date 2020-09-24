VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_DELIVERY 
   Caption         =   "DELIVERY"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   12060
   Begin OCX.b8Container b8Container1 
      Height          =   9075
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   16007
      BackColor       =   16185592
      Begin OCX.b8Container b8Container2 
         Height          =   8640
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   15240
         BorderColor     =   12735512
         BackColor       =   16185592
         Begin OCX.b8Container b8Container6 
            Height          =   7755
            Left            =   90
            TabIndex        =   6
            Top             =   120
            Width           =   11955
            _ExtentX        =   21087
            _ExtentY        =   13679
            BackColor       =   16185592
            Begin VB.Frame Frame1 
               Appearance      =   0  'Flat
               BackColor       =   &H00F6F8F8&
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   1515
               Left            =   7710
               TabIndex        =   20
               Top             =   6150
               Width           =   4185
               Begin VB.TextBox txtTotalWithoutTax 
                  Alignment       =   1  'Right Justify
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
                  Left            =   1830
                  TabIndex        =   25
                  Top             =   0
                  Width           =   2265
               End
               Begin VB.TextBox txtVAT 
                  Alignment       =   1  'Right Justify
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
                  Left            =   1830
                  TabIndex        =   24
                  Top             =   300
                  Width           =   2265
               End
               Begin VB.TextBox txtNHIL 
                  Alignment       =   1  'Right Justify
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
                  Left            =   1830
                  TabIndex        =   23
                  Top             =   600
                  Width           =   2265
               End
               Begin VB.TextBox txtTAX 
                  Alignment       =   1  'Right Justify
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
                  Left            =   1830
                  TabIndex        =   22
                  Top             =   900
                  Width           =   2265
               End
               Begin VB.TextBox txtTotalWithTax 
                  Alignment       =   1  'Right Justify
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
                  Left            =   1830
                  TabIndex        =   21
                  Top             =   1200
                  Width           =   2265
               End
               Begin VB.Label Label7 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total -TAX"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   240
                  TabIndex        =   30
                  Top             =   30
                  Width           =   1425
               End
               Begin VB.Label Label2 
                  BackStyle       =   0  'Transparent
                  Caption         =   "VAT"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   240
                  TabIndex        =   29
                  Top             =   330
                  Width           =   1545
               End
               Begin VB.Label Label5 
                  BackStyle       =   0  'Transparent
                  Caption         =   "NHIL"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   240
                  TabIndex        =   28
                  Top             =   630
                  Width           =   1425
               End
               Begin VB.Label Label6 
                  BackStyle       =   0  'Transparent
                  Caption         =   "TAX"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   240
                  TabIndex        =   27
                  Top             =   930
                  Width           =   1515
               End
               Begin VB.Label Label8 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total + TAX"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   240
                  TabIndex        =   26
                  Top             =   1230
                  Width           =   1515
               End
            End
            Begin VB.TextBox txt 
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
               Height          =   375
               Left            =   4500
               TabIndex        =   19
               Top             =   4500
               Visible         =   0   'False
               Width           =   2325
            End
            Begin OCX.b8Container ContainerCatDetails 
               Height          =   2625
               Left            =   4140
               TabIndex        =   7
               Top             =   90
               Width           =   7755
               _ExtentX        =   13679
               _ExtentY        =   4630
               BackColor       =   16185592
               Begin VB.TextBox txtContactName 
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
                  Locked          =   -1  'True
                  MaxLength       =   500
                  TabIndex        =   3
                  Top             =   1500
                  Width           =   3195
               End
               Begin VB.TextBox txtCompanyName 
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
                  Locked          =   -1  'True
                  MaxLength       =   500
                  TabIndex        =   2
                  Top             =   1050
                  Width           =   3195
               End
               Begin VB.TextBox txtOrderDate 
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
                  Locked          =   -1  'True
                  MaxLength       =   500
                  TabIndex        =   1
                  Top             =   600
                  Width           =   3195
               End
               Begin VB.TextBox txtOrderNo 
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
                  Locked          =   -1  'True
                  MaxLength       =   100
                  TabIndex        =   0
                  Top             =   150
                  Width           =   3195
               End
               Begin VB.Label Label4 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Contact Name"
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
                  TabIndex        =   14
                  Top             =   1530
                  Width           =   1365
               End
               Begin VB.Label Label3 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Company Name"
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
                  TabIndex        =   13
                  Top             =   1080
                  Width           =   1365
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
                  Left            =   180
                  TabIndex        =   9
                  Top             =   180
                  Width           =   1365
               End
               Begin VB.Label Label11 
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
                  Left            =   180
                  TabIndex        =   8
                  Top             =   630
                  Width           =   1365
               End
            End
            Begin OCX.b8Container ContainerList 
               Height          =   2625
               Left            =   60
               TabIndex        =   10
               Top             =   90
               Width           =   4035
               _ExtentX        =   7117
               _ExtentY        =   4630
               BackColor       =   16185592
               Begin VB.ListBox lstOrders 
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
                  Height          =   1980
                  Left            =   120
                  TabIndex        =   11
                  Top             =   480
                  Width           =   3765
               End
               Begin OCX.b8SideTab b8SideTab2 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   12
                  Top             =   150
                  Width           =   3765
                  _ExtentX        =   6641
                  _ExtentY        =   661
                  Caption         =   "Orders"
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
            Begin MSFlexGridLib.MSFlexGrid grd 
               Height          =   3405
               Left            =   60
               TabIndex        =   15
               Top             =   2730
               Width           =   11835
               _ExtentX        =   20876
               _ExtentY        =   6006
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
            End
         End
         Begin OCX.b8Container b8Container3 
            Height          =   735
            Left            =   90
            TabIndex        =   16
            Top             =   7860
            Width           =   11955
            _ExtentX        =   21087
            _ExtentY        =   1296
            BackColor       =   16185592
            Begin lvButton.lvButtons_H cmdClose 
               Height          =   495
               Left            =   10350
               TabIndex        =   17
               Top             =   120
               Width           =   1335
               _ExtentX        =   2355
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
               Image           =   "frm_DELIVERY.frx":0000
               ImgSize         =   24
               cBack           =   -2147483633
               mPointer        =   99
               mIcon           =   "frm_DELIVERY.frx":0590
            End
            Begin lvButton.lvButtons_H cmdSave 
               Height          =   495
               Left            =   8850
               TabIndex        =   18
               Top             =   120
               Width           =   1335
               _ExtentX        =   2355
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
               Image           =   "frm_DELIVERY.frx":08AA
               ImgSize         =   24
               cBack           =   -2147483633
               mPointer        =   99
               mIcon           =   "frm_DELIVERY.frx":0B0A
            End
         End
      End
   End
End
Attribute VB_Name = "frm_DELIVERY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngOrderID As Long
Dim lngProductID As Long
Dim lngSupplierID As Long
Dim lngDeliveryID As Long
Dim strProductName As String
Dim m_row, m_col As String

Private Sub sub_LOAD_ORDERS(Optional lst As ListBox)

    Dim rec As New ADODB.Recordset
    Set rec = cls_DELIVERY_Obj.fn_LOAD_ORDERS(0)
    
    lst.Clear
    
    Do While Not rec.EOF
        lst.AddItem rec!OrderNo
        lst.ItemData(lst.NewIndex) = rec!OrderID
        rec.MoveNext
    Loop

End Sub

Private Sub sub_LOAD_ORDERS_DETAILS(lngOrderID As Long)

    Dim rec As New ADODB.Recordset
    Set rec = cls_DELIVERY_Obj.fn_LOAD_ORDERS(lngOrderID)
    
    If rec.AbsolutePosition <> -1 Then
        txtOrderNo.Text = rec!OrderNo
        txtOrderDate.Text = rec!OrderDate
        Call sub_LOAD_SUPPLIERS_DETAILS(rec!SupplierID)
    End If

End Sub



Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim lngStatementID As Long
    Dim lngRow As Long
    If grd.Rows <= 2 Then Exit Sub
    If MsgBox("Are you sure you want to Save ?", vbYesNo + vbQuestion, Title) = vbYes Then
        
        With cls_DELIVERY_Obj
            lngDeliveryID = .fn_AUTOGEN
            .DeliveryID = lngDeliveryID
            .OrderID = lngOrderID
            .DeliveryDate = Date
            .DeliveryTime = Now
            .TotalWithoutTax = Val(txtTotalWithoutTax.Text)
            .TAX = Val(txtTAX.Text)
            .TotalWithTax = Val(txtTotalWithTax.Text)
            .fn_SAVE_DELIVERY_RECORDS
        End With
          
        For lngRow = 1 To grd.Rows - 2
            With cls_DELIVERY_Obj
                .DeliveryID = lngDeliveryID
                .ProductID = grd.TextMatrix(lngRow, 6)
                .SupplierPrice = grd.TextMatrix(lngRow, 3)
                .UnitsOrdered = grd.TextMatrix(lngRow, 4)
                .UnitsReceived = grd.TextMatrix(lngRow, 5)
                .VAT = Val(grd.TextMatrix(lngRow, 3)) * Val(grd.TextMatrix(lngRow, 4)) * lngVAT / 100
                .NHIL = Val(grd.TextMatrix(lngRow, 3)) * Val(grd.TextMatrix(lngRow, 4)) * lngNHIL / 100
                .fn_SAVE_DELIVERY_DETAILS_RECORDS
                .fn_UPDATE_PRODUCTS_IN_STOCK (.ProductID)
                .fn_UPDATE_ORDERS_STATUS (lngOrderID)
            End With
        Next
        
        MsgBox "Transaction Done successfully", vbInformation, Title
    
        Call sub_FORMAT_GRID
        Call sub_LOAD_ORDERS(lstOrders)
        
    End If
    
    
EXITPROCEDURE:
    Exit Sub
End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_FORM_SIZE(Me)
    Move 0, 0
    Call sub_LOAD_ORDERS(lstOrders)
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSave, cmdSave.hwnd, "Save delivery details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdClose, cmdClose.hwnd, "Cancel and close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor

End Sub

Private Sub grd_Click()
    
    If grd.col = 5 Then
        m_row = grd.Row
        m_col = grd.col
        sub_GET_POSITION
        Call Mdl_FUNCTIONS.fn_HIGHLIGHT_TEXT(txt)
    End If
    
End Sub

Private Sub lstOrders_Click()
    If lstOrders.ListIndex = -1 Then Exit Sub
    lngOrderID = lstOrders.ItemData(lstOrders.ListIndex)
    Call sub_LOAD_ORDERS_DETAILS(lngOrderID)
    Call sub_LOAD_ORDER_PRODUCTS(lngOrderID)
End Sub

Private Sub sub_LOAD_SUPPLIERS_DETAILS(lngSupplierID As Long)

    Dim rec As New ADODB.Recordset
    
    Set rec = cls_SUPPLIER_Obj.fn_LOAD_SUPPLIERS(lngSupplierID)
    
    
    With ClsSupplierObject
    
        txtCompanyName.Text = rec!CompanyName
        txtContactName.Text = rec!ContactName
        
    End With
    
End Sub
Private Sub sub_LOAD_ORDER_PRODUCTS(lngOrderID As Long)

    Dim rec As New ADODB.Recordset
    Set rec = cls_ORDER_Obj.fn_LOAD_ORDERS_DETAILS(lngOrderID)
    
    Call sub_FORMAT_GRID
    
    Do While Not rec.EOF
        With grd
            .TextMatrix(.Rows - 1, 1) = rec!OrderDetailsID
            Call sub_LOAD_PRODUCTS_DETAILS(rec!ProductID)
            .TextMatrix(.Rows - 1, 2) = strProductName
            .TextMatrix(.Rows - 1, 3) = rec!SupplierPrice
            .TextMatrix(.Rows - 1, 4) = rec!Qty
            .TextMatrix(.Rows - 1, 5) = rec!Qty
            .TextMatrix(.Rows - 1, 6) = rec!ProductID
            .TextMatrix(.Rows - 1, 7) = Val(rec!SupplierPrice) * Val(rec!Qty)
            txtTotalWithoutTax.Text = Val(txtTotalWithoutTax.Text) + Val(.TextMatrix(.Rows - 1, 7))
            .Rows = .Rows + 1
        End With
        rec.MoveNext
    Loop
    
    txtVat.Text = Val(txtTotalWithoutTax.Text) * lngVAT / 100
    txtNHIL.Text = Val(txtTotalWithoutTax.Text) * lngNHIL / 100
    txtTAX.Text = Val(txtVat.Text) + Val(txtNHIL.Text)
    txtTotalWithTax.Text = Val(txtTotalWithoutTax) + Val(txtTAX.Text)
    
End Sub

Private Sub sub_FORMAT_GRID()

    With grd
        .Clear
        .FormatString = "|Order Details No|Product|>Supplier Price|>Units Ordered|>Units Received|ProductID|Total"
        .ColWidth(0) = 300
        .ColWidth(1) = 1800
        .ColWidth(2) = 4000
        .ColWidth(3) = 1850
        .ColWidth(4) = 1850
        .ColWidth(5) = 1850
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .Rows = 2
    End With
    txtTotalWithoutTax.Text = ""
    txtVat.Text = ""
    txtNHIL.Text = ""
    txtTAX.Text = ""
    txtTotalWithTax.Text = ""
End Sub

Private Sub sub_LOAD_PRODUCTS_DETAILS(lngProductID As Long)
    Dim rec As New ADODB.Recordset

    Set rec = cls_PRODUCT_Obj.fN_LOAD_PRODUCTS_DETAILS(lngProductID)

    If rec.AbsolutePosition <> -1 Then
        strProductName = Trim(rec!ProductName) & ""
    End If

End Sub

Private Sub sub_GET_POSITION()
    grd.Row = m_row
    grd.col = m_col
    Call sub_SIZE
End Sub
Private Sub sub_SIZE()

    txt = grd.TextMatrix(m_row, m_col)
    txt.Left = grd.Left + grd.CellLeft
    txt.Top = grd.Top + grd.CellTop
    txt.Width = grd.CellWidth
    txt.Height = grd.CellHeight
    txt.Visible = True
    
End Sub
Private Sub txt_Change()
    grd.TextMatrix(m_row, m_col) = txt
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 39 And m_col < grd.Cols - 1 Then m_col = m_col + 1
    If KeyCode = 37 And m_col > 1 Then m_col = m_col - 1
    If KeyCode = 40 And m_row < grd.Rows - 1 Then m_row = m_row + 1
    If KeyCode = 38 And m_row > 1 Then m_row = m_row - 1
    If KeyCode = 39 Or KeyCode = 37 Or KeyCode = 40 Or KeyCode = 38 Then sub_GET_POSITION
End Sub
