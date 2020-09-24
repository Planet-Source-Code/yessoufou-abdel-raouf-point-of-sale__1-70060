VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm frm_MAIN 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   1155
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3990
      Top             =   4110
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4680
      TabIndex        =   4
      Top             =   2475
      Width           =   4680
      Begin OCX.b8Container ctn2 
         Height          =   615
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   1085
         BorderColor     =   16711680
         BackColor       =   16777215
         Begin VB.Label lblTime 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   11325
            TabIndex        =   11
            Top             =   180
            Width           =   3765
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   5010
            TabIndex        =   10
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label lblRole 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   915
            TabIndex        =   9
            Top             =   210
            Width           =   2985
         End
         Begin VB.Image Image3 
            Height          =   480
            Left            =   60
            Picture         =   "frm_MAIN.frx":0000
            Stretch         =   -1  'True
            Top             =   60
            Width           =   690
         End
         Begin VB.Image Image4 
            Height          =   660
            Left            =   4260
            Picture         =   "frm_MAIN.frx":5C12
            Stretch         =   -1  'True
            Top             =   0
            Width           =   750
         End
         Begin VB.Label lblName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   6540
            TabIndex        =   8
            Top             =   180
            Width           =   2565
         End
         Begin VB.Image Image5 
            Height          =   480
            Left            =   9180
            Picture         =   "frm_MAIN.frx":78DC
            Top             =   90
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date && Time:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   9660
            TabIndex        =   7
            Top             =   180
            Width           =   1605
         End
         Begin VB.Image imgBotom 
            Height          =   495
            Left            =   60
            Picture         =   "frm_MAIN.frx":7D1E
            Stretch         =   -1  'True
            Top             =   60
            Width           =   15150
         End
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4680
      TabIndex        =   2
      Top             =   0
      Width           =   4680
      Begin OCX.b8Container ctn1 
         Height          =   615
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   1085
         BorderColor     =   16711680
         BackColor       =   16777215
         Begin VB.Image imgTop 
            Height          =   495
            Left            =   60
            Picture         =   "frm_MAIN.frx":9FAD
            Stretch         =   -1  'True
            Top             =   60
            Width           =   15150
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   0
      ScaleHeight     =   1860
      ScaleWidth      =   3120
      TabIndex        =   0
      Top             =   615
      Width           =   3120
      Begin OCX.b8SideTab SideBar1 
         Height          =   10575
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   18653
         Caption         =   "Menu"
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
         Begin MSComctlLib.TreeView tvw 
            Height          =   8895
            Left            =   30
            TabIndex        =   6
            Top             =   480
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   15690
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   840
      Top             =   1680
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuChangeUserName 
         Caption         =   "Change User Name && Password"
      End
      Begin VB.Menu s06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "Log Off"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLockApplication 
         Caption         =   "Lock Application"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuitApplication 
         Caption         =   "Quit Application"
      End
   End
   Begin VB.Menu mnuTansactions 
      Caption         =   "&Transactions"
      Begin VB.Menu mnuSales 
         Caption         =   "Sales"
         Begin VB.Menu mnuCashSales 
            Caption         =   "Cash Sales"
         End
         Begin VB.Menu s070 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCreditSales 
            Caption         =   "Credit Sales"
         End
         Begin VB.Menu s034 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAllCreditSales 
            Caption         =   "All Credit Sales"
         End
      End
      Begin VB.Menu s4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrders 
         Caption         =   "Orders"
         Begin VB.Menu mnuOrdersToSuppliers 
            Caption         =   "Orders To Suppliers"
         End
         Begin VB.Menu s031 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuOrdersFromCustomers 
            Caption         =   "Orders From Customers"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelivery 
         Caption         =   "Delivery"
      End
      Begin VB.Menu s014 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExpenditures 
         Caption         =   "Expenditures"
      End
      Begin VB.Menu s030 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustomersDeposit 
         Caption         =   "Customers Deposit"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuEmployeesRep 
         Caption         =   "Employees"
         Begin VB.Menu mnuViewEmployeesRep 
            Caption         =   "View Employees Reports"
         End
      End
      Begin VB.Menu s0000 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustomersReport 
         Caption         =   "Customers"
         Begin VB.Menu mnuViewCustomersReport 
            Caption         =   "View Customers Report"
         End
      End
      Begin VB.Menu s001 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSuppliersReport 
         Caption         =   "Suppliers"
         Begin VB.Menu mnuViewSuppliersAndProduct 
            Caption         =   "View Suppliers And Products"
         End
      End
      Begin VB.Menu s012 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExpendituresReport 
         Caption         =   "Expenditures"
         Begin VB.Menu mnuViewExpendituresReport 
            Caption         =   "View Expenditures Report"
         End
      End
      Begin VB.Menu s020 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalesReports 
         Caption         =   "Sales"
         Begin VB.Menu mnuViewSalesReports 
            Caption         =   "View Sales Report"
         End
         Begin VB.Menu mnuTaxReport 
            Caption         =   "View Tax Report"
         End
      End
      Begin VB.Menu s6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrdersReports 
         Caption         =   "Orders"
         Begin VB.Menu mnuViewOrdersReports 
            Caption         =   "View Orders Report"
         End
         Begin VB.Menu mnuPendingOrdersReport 
            Caption         =   "View Pending Orders Report"
         End
         Begin VB.Menu mnuOrdersNotPending 
            Caption         =   "View Orders Not Pending"
         End
      End
      Begin VB.Menu s7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeliveryReports 
         Caption         =   "Delivery"
         Begin VB.Menu mnuViewDeliveryReports 
            Caption         =   "View Delivery Report"
         End
         Begin VB.Menu mnuViewDeliveryTaxReport 
            Caption         =   "View Delivery Tax Report"
         End
      End
      Begin VB.Menu s01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStockReport 
         Caption         =   "Stock"
         Begin VB.Menu mnuViewStockReport 
            Caption         =   "View Stock Report"
         End
         Begin VB.Menu s05 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewProductsPrice 
            Caption         =   "View Products Price"
         End
         Begin VB.Menu s066 
            Caption         =   "-"
         End
         Begin VB.Menu mnuActiveProducts 
            Caption         =   "View Active Products"
         End
         Begin VB.Menu s07 
            Caption         =   "-"
         End
         Begin VB.Menu mnuInactiveProducts 
            Caption         =   "View Inactive Products"
         End
         Begin VB.Menu s08 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewProductsOutOfStock 
            Caption         =   "View Products Out Of Stock"
         End
         Begin VB.Menu s09 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewProductsToReorder 
            Caption         =   "View Products To Be Reordered"
         End
      End
      Begin VB.Menu s02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProfitAndLoss 
         Caption         =   "Profit && Loss"
         Begin VB.Menu mnuViewProfitAndLoss 
            Caption         =   "View Profit && Loss"
         End
      End
      Begin VB.Menu s21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBankReport 
         Caption         =   "Bank"
         Begin VB.Menu mnuViewBankTransactions 
            Caption         =   "View Bank Transactions"
         End
      End
      Begin VB.Menu s000 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsersReports 
         Caption         =   "Users"
         Begin VB.Menu mnuViewUsersAccessLogs 
            Caption         =   "View Users Access Logs"
         End
      End
   End
   Begin VB.Menu mnuAdministration 
      Caption         =   "&Administration"
      Begin VB.Menu mnuCompanyInfo 
         Caption         =   "Company Info"
         Begin VB.Menu mnuViewCompanyInfo 
            Caption         =   "View Company Info"
         End
      End
      Begin VB.Menu s005 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmployees 
         Caption         =   "Employees"
         Begin VB.Menu mnuViewAllEmployees 
            Caption         =   "View All Employees"
         End
         Begin VB.Menu s006 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLeaves 
            Caption         =   "View Leaves"
         End
         Begin VB.Menu s007 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSalaries 
            Caption         =   "View Salaries"
         End
         Begin VB.Menu s008 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewSalaryPeriod 
            Caption         =   "View Salary Period"
         End
      End
      Begin VB.Menu s04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustomers 
         Caption         =   "Customers"
         Begin VB.Menu mnuViewAllCustomers 
            Caption         =   "View All Customers"
         End
      End
      Begin VB.Menu s10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSuppliers 
         Caption         =   "Suppliers"
         Begin VB.Menu mnuAllSuppliers 
            Caption         =   "View All Suppliers"
         End
      End
      Begin VB.Menu s0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCategories 
         Caption         =   "Categories"
         Begin VB.Menu mnuViewAllCategories 
            Caption         =   "View All Categories"
         End
      End
      Begin VB.Menu s11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProducts 
         Caption         =   "Products"
         Begin VB.Menu mnuViewAllProducts 
            Caption         =   "View All Products"
         End
      End
      Begin VB.Menu s12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Users"
         Begin VB.Menu mnuViewAllUsers 
            Caption         =   "View All Users"
         End
      End
      Begin VB.Menu S032 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBanks 
         Caption         =   "Banks"
         Begin VB.Menu mnuBankAndAccount 
            Caption         =   "Banks && Accounts"
         End
         Begin VB.Menu mnuBankDeposit 
            Caption         =   "Bank Deposit"
         End
         Begin VB.Menu mnuBankWithdrawal 
            Caption         =   "Bank Withdrawal"
         End
         Begin VB.Menu mnuBankCharges 
            Caption         =   "Bank Charges"
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCalculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu s011 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNotepad 
         Caption         =   "Notepad"
      End
      Begin VB.Menu s13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuServerConnection 
         Caption         =   "Server Connection"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      NegotiatePosition=   1  'Left
      WindowList      =   -1  'True
      Begin VB.Menu mnuCloseAllOpenedWindows 
         Caption         =   "Close All Opened Windows"
      End
   End
End
Attribute VB_Name = "frm_MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lngCurrentUserID = 0 Then Exit Sub
    Call frm_LOGIN.sub_SAVE_USERS_LOGS(0, "Log Off")
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If lngCurrentUserID = 0 Then Exit Sub
    Call frm_LOGIN.sub_SAVE_USERS_LOGS(0, "Log Off")
End Sub

Private Sub mnuActiveProducts_Click()
    If mnuActiveProducts.Enabled = False Then Exit Sub
    frm_ACTIVE_PRODUCTS.Show
End Sub

Private Sub mnuAllCreditSales_Click()
    If mnuAllCreditSales.Enabled = False Then Exit Sub
    With frm_PENDING_SALES
'        lblCaption.Caption ="ALL CREDIT SALES"
        .SetFocus
        .Show
    End With
End Sub

Private Sub mnuAllSuppliers_Click()
    If mnuAllSuppliers.Enabled = False Then Exit Sub
    frm_SUPPLIERS.SetFocus
    frm_SUPPLIERS.Show
End Sub

Private Sub mnuBankAndAccount_Click()
    If mnuBankAndAccount.Enabled = False Then Exit Sub
    With frm_BANKS
        .SetFocus
        .Show
    End With
End Sub

Private Sub mnuBankCharges_Click()
    If mnuBankCharges.Enabled = False Then Exit Sub
    frm_BANK_CHARGES.Show 1
End Sub

Private Sub mnuBankDeposit_Click()
    If mnuBankDeposit.Enabled = False Then Exit Sub
    With frm_DEPOSIT
        .Show 1
    End With
End Sub

Private Sub mnuBankWithdrawal_Click()
    If mnuBankWithdrawal.Enabled = False Then Exit Sub
    With frm_WITHDRAWAL
        .Show 1
    End With
End Sub

Private Sub mnuCashSales_Click()
    If mnuCashSales.Enabled = False Then Exit Sub
    With frm_SALES
'        lblCaption.Caption = "CASH SALES"
        .cmdSave.Enabled = True
        .cmdHold.Enabled = False
        .SetFocus
        .Show
    End With
End Sub

Private Sub mnuChangeUserName_Click()
    frm_USER_NAME.Show
End Sub

Private Sub mnuCreditSales_Click()
    If mnuCreditSales.Enabled = False Then Exit Sub
    With frm_SALES
'        lblCaption.Caption ="CREDIT SALES"
        .cmdHold.Enabled = True
        .cmdSave.Enabled = False
        .SetFocus
        .Show
    End With
End Sub

Private Sub mnuCustomersDeposit_Click()
    If mnuCustomersDeposit.Enabled = False Then Exit Sub
    With frm_CUSTOMER_DEPOSIT
'        .SetFocus
        .Show
    End With
End Sub

Private Sub mnuDelivery_Click()
    If mnuDelivery.Enabled = False Then Exit Sub
    frm_DELIVERY.SetFocus
    frm_DELIVERY.Show
End Sub

Private Sub mnuExpenditures_Click()
    If mnuExpenditures.Enabled = False Then Exit Sub
    frm_EXPENDITURES.Show
End Sub

Private Sub mnuInactiveProducts_Click()
    If mnuInactiveProducts.Enabled = False Then Exit Sub
    frm_INACTIVE_PRODUCTS.Show
End Sub

Private Sub mnuLeaves_Click()
    With frm_LEAVES
        .SetFocus
        .Show
    End With
End Sub

'*******File Menu*****************************************************************
    
Private Sub mnuLogOff_Click()
    If MsgBox("All unsaved transactions will be lost." & vbCrLf & "Are you sure you want to log out?", vbYesNo, Title) = vbNo Then
        GoTo EXITPROCEDURE
        Else
            Call frm_LOGIN.sub_SAVE_USERS_LOGS(0, "Log Off")
            mnuCloseallOpenedForms_Click
            frm_MAIN.Show
            frm_LOGIN.Show 1
    End If
    
EXITPROCEDURE:
    Exit Sub
End Sub

Private Sub mnuLockApplication_Click()
    frm_LOGIN.Show 1
End Sub

Private Sub mnuNewSales_Click()
    If mnuNewSales.Enabled = False Then Exit Sub
    With frm_SALES
        .SetFocus
        .Show
    End With
End Sub


Private Sub mnuOrdersFromCustomers_Click()
    If mnuOrdersFromCustomers.Enabled = False Then Exit Sub
    With frm_ORDERS_FROM_CUSTOMERS
        .SetFocus
        .Show
    End With
End Sub

Private Sub mnuOrdersNotPending_Click()
    If mnuOrdersNotPending.Enabled = False Then Exit Sub
    frm_ORDERS_NOT_PENDING.Show
End Sub

Private Sub mnuOrdersToSuppliers_Click()
    If mnuOrdersToSuppliers.Enabled = False Then Exit Sub
    frm_ORDERS.SetFocus
    frm_ORDERS.Show
End Sub

Private Sub mnuPendingOrdersReport_Click()
    If mnuPendingOrdersReport.Enabled = False Then Exit Sub
    frm_PENDING_ORDERS_REP.Show
End Sub

Private Sub mnuPendingSales_Click()
    If mnuPendingSales.Enabled = False Then Exit Sub
    With frm_SALES
        .SetFocus
        .Show
    End With
End Sub

Private Sub mnuQuitApplication_Click()
    If MsgBox("Are you sure you want to quit the application?", 4 + 32, Title) = vbYes Then
        Call frm_LOGIN.sub_SAVE_USERS_LOGS(0, "Log Off")
        End
    End If
End Sub




Private Sub mnuSalaries_Click()
    With frm_SALARIES
        .SetFocus
        .Show
    End With
End Sub

Private Sub mnuServerConnection_Click()
    If mnuServerConnection.Enabled = False Then Exit Sub
    With frm_SERVER_CONNECTION
        .Show 1
    End With
End Sub

Private Sub mnuTaxReport_Click()
    If mnuTaxReport.Enabled = False Then Exit Sub
    frm_TAX_REP.Show
End Sub

'*********************************************************************************


'*******Transactions Menu***************************************************************



'*********************************************************************************


'*******Reports Menu***************************************************************



'*********************************************************************************


'*******Administration Menu***************************************************************


Private Sub mnuViewAllCategories_Click()
    If mnuViewAllCategories.Enabled = False Then Exit Sub
    With frm_CATEGORIES
        .SetFocus
        .Show
    End With
End Sub

Private Sub mnuViewAllCustomers_Click()
    If mnuViewAllCustomers.Enabled = False Then Exit Sub
    frm_CUSTOMERS.SetFocus
    frm_CUSTOMERS.Show
End Sub

Private Sub mnuViewAllEmployees_Click()
    With frm_EMPLOYEES
        .SetFocus
        .Show
    End With
End Sub

Private Sub mnuViewAllProducts_Click()
    If mnuViewAllProducts.Enabled = False Then Exit Sub
    frm_PRODUCTS.SetFocus
    frm_PRODUCTS.Show
End Sub


'*********************************************************************************


'*******Tools Menu***************************************************************
    
    Private Sub mnunotepad_Click()
        Shell "Notepad.exe", vbNormalFocus
    End Sub
    
    Private Sub mnuCalculator_Click()
        Shell "calc"
    End Sub

'********************************************************************************


'*******Windows Menu***************************************************************

    Private Sub mnuCloseallOpenedForms_Click()
        Call Mdl_FUNCTIONS.sub_CLOSE_ALL_OPENED_FORMS
    End Sub

'*********************************************************************************


Private Sub MDIForm_Load()
    Call Me.SubSetTreeView
    Call sub_DISABLE_FEATURES
End Sub

Private Sub MDIForm_Resize()

    ctn1.Width = frm_MAIN.Width
    imgTop.Width = ctn1.Width
    ctn2.Width = frm_MAIN.Width
    imgBotom.Width = ctn2.Width
    SideBar1.Width = Picture1.Width

End Sub



Private Sub mnuViewAllUsers_Click()
    If mnuViewAllUsers.Enabled = False Then Exit Sub
    frm_USERS.SetFocus
    frm_USERS.Show
End Sub

Private Sub mnuViewBankTransactions_Click()
    If mnuViewBankTransactions.Enabled = False Then Exit Sub
    With frm_BANK_TRANSACTIONS
        .SetFocus
        .Show
    End With
End Sub

Private Sub mnuViewCompanyInfo_Click()
    If mnuViewCompanyInfo.Enabled = False Then Exit Sub
    frm_COMPANY_INFO.Show 1
End Sub

Private Sub mnuViewCustomersReport_Click()
    If mnuViewCustomersReport.Enabled = False Then Exit Sub
    frm_CUSTOMERS_REPORT.Show
End Sub

Private Sub mnuViewDeliveryReports_Click()
    If mnuViewDeliveryReports.Enabled = False Then Exit Sub
    frm_DELIVERY_REP.Show
End Sub

Private Sub mnuViewDeliveryTaxReport_Click()
    If mnuViewDeliveryTaxReport.Enabled = False Then Exit Sub
    With frm_DELIVERY_TAX_REP
        .Show
    End With
End Sub

Private Sub mnuViewEmployeesRep_Click()
    If mnuViewAllEmployees.Enabled = False Then Exit Sub
    frm_EMPLOYEES_REP.Show
End Sub

Private Sub mnuViewExpendituresReport_Click()
    If mnuViewExpendituresReport.Enabled = False Then Exit Sub
    frm_EXPENDITURES_REPPORT.Show
End Sub

Private Sub mnuViewOrdersReports_Click()
    If mnuViewOrdersReports.Enabled = False Then Exit Sub
    frm_ORDERS_REP.Show
End Sub

Private Sub mnuViewProductsOutOfStock_Click()
    If mnuViewProductsOutOfStock.Enabled = False Then Exit Sub
    frm_PRODUCTS_OUT_OF_STOCK_REP.Show
End Sub

Private Sub mnuViewProductsPrice_Click()
    If mnuViewProductsPrice.Enabled = False Then Exit Sub
    frm_STOCK_REP.Show
End Sub

Private Sub mnuViewProductsToReorder_Click()
    If mnuViewProductsToReorder.Enabled = False Then Exit Sub
    frm_PRODUCT_TO_REORDER_REP.Show
End Sub

Private Sub mnuViewProfitAndLoss_Click()
    If mnuViewProfitAndLoss.Enabled = False Then Exit Sub
    frm_PROFIT_LOSS.Show
End Sub

Private Sub mnuViewSalaryPeriod_Click()
    If mnuViewSalaryPeriod.Enabled = False Then Exit Sub
    With frm_SALARY_PERIOD
        .Show
    End With
End Sub

Private Sub mnuViewSalesReports_Click()
    If mnuViewSalesReports.Enabled = False Then Exit Sub
    frm_SALES_REP.Show
End Sub

Private Sub mnuViewStockReport_Click()
    If mnuViewStockReport.Enabled = False Then Exit Sub
    frm_STOCK_CUMULATIVE_REP.Show
End Sub

Private Sub mnuViewSuppliersAndProduct_Click()
    If mnuViewSuppliersAndProduct.Enabled = False Then Exit Sub
    frm_SUPPLIERS_AND_PRODUCTS.Show
End Sub

Private Sub mnuViewUsersAccessLogs_Click()
    If mnuViewUsersAccessLogs.Enabled = False Then Exit Sub
    frm_USERS_ACCESS_LOGS.Show
End Sub

Private Sub Timer1_Timer()
    Caption = Left(msgToScroll, curPosition)
    curPosition = (curPosition + 1) Mod (Len(msgToScroll) + 1)
End Sub


Public Sub SubSetTreeView()

    Set XNode = tvw.Nodes.Add(, , "mnuFile", "File")
    Call Mdl_FUNCTIONS.sub_SET_TREEVIEW_FONT
    Set XNode = tvw.Nodes.Add("mnuFile", tvwChild, "mnuChangeUserName", "Change User Name & Password")
    Set XNode = tvw.Nodes.Add("mnuFile", tvwChild, "mnuLogOff", "Log Off")
    Set XNode = tvw.Nodes.Add("mnuFile", tvwChild, "mnuLockApplication", "Lock Application")
    Set XNode = tvw.Nodes.Add("mnuFile", tvwChild, "mnuQuitApplication", "Quit Application")

    Set XNode = tvw.Nodes.Add(, , "mnuTransactions", "Transactions")
    Call Mdl_FUNCTIONS.sub_SET_TREEVIEW_FONT
    Set XNode = tvw.Nodes.Add("mnuTransactions", tvwChild, "mnuSales", "Sales")
        Set XNode = tvw.Nodes.Add("mnuSales", tvwChild, "mnuCashSales", "Cash Sales")
        Set XNode = tvw.Nodes.Add("mnuSales", tvwChild, "mnuCreditSales", "Credit Sales")
        Set XNode = tvw.Nodes.Add("mnuSales", tvwChild, "mnuAllSales", "All Credit Sales")
    Set XNode = tvw.Nodes.Add("mnuTransactions", tvwChild, "mnuOrders", "Orders")
        Set XNode = tvw.Nodes.Add("mnuOrders", tvwChild, "mnuOrdersToSuppliers", "Orders To Suppliers")
'        Set XNode = tvw.Nodes.Add("mnuOrders", tvwChild, "mnuOrdersFromCustomers", "Orders From Customers")
                
    Set XNode = tvw.Nodes.Add("mnuTransactions", tvwChild, "mnuDelivery", "Delivery")
    Set XNode = tvw.Nodes.Add("mnuTransactions", tvwChild, "mnuExpenditures", "Expenditures")
    Set XNode = tvw.Nodes.Add("mnuTransactions", tvwChild, "mnuCustomersDeposit", "Customers Deposit")
    
    Set XNode = tvw.Nodes.Add(, , "mnuReports", "Reports")
    Call Mdl_FUNCTIONS.sub_SET_TREEVIEW_FONT
    Set XNode = tvw.Nodes.Add("mnuReports", tvwChild, "mnuEmployeesReport", "Employees")
    Set XNode = tvw.Nodes.Add("mnuReports", tvwChild, "mnuCustomersReport", "Customers")
    Set XNode = tvw.Nodes.Add("mnuReports", tvwChild, "mnuSuppliersReport", "Suppliers")
    Set XNode = tvw.Nodes.Add("mnuReports", tvwChild, "mnuExpendituresReport", "Expenditures")
    Set XNode = tvw.Nodes.Add("mnuReports", tvwChild, "mnuSalesReport", "Sales")
        Set XNode = tvw.Nodes.Add("mnuSalesReport", tvwChild, "mnuViewSalesReport", "Sales Report")
        Set XNode = tvw.Nodes.Add("mnuSalesReport", tvwChild, "mnuTaxReport", "Tax Report")
    Set XNode = tvw.Nodes.Add("mnuReports", tvwChild, "mnuOrdersReports", "Orders")
        Set XNode = tvw.Nodes.Add("mnuOrdersReports", tvwChild, "mnuViewOrdersReports", "View Orders Reports")
        Set XNode = tvw.Nodes.Add("mnuOrdersReports", tvwChild, "mnuPendingOrdersReport", "View Orders Pending")
        Set XNode = tvw.Nodes.Add("mnuOrdersReports", tvwChild, "mnuOrdersNotPending", "View Orders Not Pending")
    Set XNode = tvw.Nodes.Add("mnuReports", tvwChild, "mnuDeliveryReports", "Delivery")
        Set XNode = tvw.Nodes.Add("mnuDeliveryReports", tvwChild, "mnuViewDeliveryReports", "View Delivery Report")
        Set XNode = tvw.Nodes.Add("mnuDeliveryReports", tvwChild, "mnuViewDeliveryTaxReport", "View Delivery Tax Report")
     Set XNode = tvw.Nodes.Add("mnuReports", tvwChild, "mnuStock", "Stock")
        Set XNode = tvw.Nodes.Add("mnuStock", tvwChild, "mnuViewStockReport", "View Stock Report")
        Set XNode = tvw.Nodes.Add("mnuStock", tvwChild, "mnuViewProductsPrice", "View Products Price")
        Set XNode = tvw.Nodes.Add("mnuStock", tvwChild, "mnuActiveProducts", "View Active Products")
        Set XNode = tvw.Nodes.Add("mnuStock", tvwChild, "mnuInactiveProducts", "View Inactive Products")
        Set XNode = tvw.Nodes.Add("mnuStock", tvwChild, "mnuViewProductsOutOfStock", "View Products Out Of Stock")
        Set XNode = tvw.Nodes.Add("mnuStock", tvwChild, "mnuViewProductsToReorder", "View Products To Reorder")
    Set XNode = tvw.Nodes.Add("mnuReports", tvwChild, "mnuProfitLoss", "Profit and Loss")
        Set XNode = tvw.Nodes.Add("mnuProfitLoss", tvwChild, "mnuViewProfitLoss", "View Profit and Loss")
    
    Set XNode = tvw.Nodes.Add("mnuReports", tvwChild, "mnuBankTransaction", "Banks Transactions")
    Set XNode = tvw.Nodes.Add("mnuReports", tvwChild, "mnuUsersTimeRecords", "User Time Records")
        Set XNode = tvw.Nodes.Add("mnuUsersTimeRecords", tvwChild, "mnuViewUsersTimeRecords", "View Users Time Records")
    
    
    Set XNode = tvw.Nodes.Add(, , "mnuAdministration", "Administration")
    Call Mdl_FUNCTIONS.sub_SET_TREEVIEW_FONT
    Set XNode = tvw.Nodes.Add("mnuAdministration", tvwChild, "mnuCompanyInfo", "CompanyInfo")
    Set XNode = tvw.Nodes.Add("mnuAdministration", tvwChild, "mnuEmployees", "Employees")
        Set XNode = tvw.Nodes.Add("mnuEmployees", tvwChild, "mnuViewEmployees", "View Employees")
        Set XNode = tvw.Nodes.Add("mnuEmployees", tvwChild, "mnuViewLeaves", "View Leaves")
        Set XNode = tvw.Nodes.Add("mnuEmployees", tvwChild, "mnuViewSalaries", "View Salaries")
        Set XNode = tvw.Nodes.Add("mnuEmployees", tvwChild, "mnuViewSalaryPeriod", "View Salary Period")
    Set XNode = tvw.Nodes.Add("mnuAdministration", tvwChild, "mnuCustomers", "Customers")
    Set XNode = tvw.Nodes.Add("mnuAdministration", tvwChild, "mnuSuppliers", "Suppliers")
    Set XNode = tvw.Nodes.Add("mnuAdministration", tvwChild, "mnuCategories", "Categories")
    Set XNode = tvw.Nodes.Add("mnuAdministration", tvwChild, "mnuProducts", "Products")
    Set XNode = tvw.Nodes.Add("mnuAdministration", tvwChild, "mnuUsers", "Users")
    Set XNode = tvw.Nodes.Add("mnuAdministration", tvwChild, "mnuBanks", "Banks")
        Set XNode = tvw.Nodes.Add("mnuBanks", tvwChild, "mnuBankAndAccount", "Banks & Accounts")
        Set XNode = tvw.Nodes.Add("mnuBanks", tvwChild, "mnuDeposit", "Deposit")
        Set XNode = tvw.Nodes.Add("mnuBanks", tvwChild, "mnuWithdrawal", "Withdrawal")
        Set XNode = tvw.Nodes.Add("mnuBanks", tvwChild, "mnuCharges", "Charges")
    Set XNode = tvw.Nodes.Add(, , "mnuTools", "Tools")
    Call Mdl_FUNCTIONS.sub_SET_TREEVIEW_FONT
    Set XNode = tvw.Nodes.Add("mnuTools", tvwChild, "mnuCalculator", "Calculator")
    Set XNode = tvw.Nodes.Add("mnuTools", tvwChild, "mnuNotepad", "Notepad")


End Sub

Private Sub sub_LOAD_FORMS(strNodeKey As String)

    Select Case strNodeKey
    
        Case "mnuChangeUserName"
            Call mnuChangeUserName_Click
            
        Case "mnuLogOff"
            Call mnuLogOff_Click
        
        Case "mnuLockApplication"
            Call mnuLockApplication_Click

        Case "mnuQuitApplication"
            Call mnuQuitApplication_Click
            
            
        Case "mnuCashSales"
            Call mnuCashSales_Click
            
        Case "mnuCreditSales"
            Call mnuCreditSales_Click
            
        Case "mnuAllSales"
            Call mnuAllCreditSales_Click


        Case "mnuOrdersToSuppliers"
            Call mnuOrdersToSuppliers_Click
            
'        Case "mnuOrdersFromCustomers"
'            Call mnuOrdersFromCustomers_Click
            
        Case "mnuDelivery"
            Call mnuDelivery_Click
            
        Case "mnuExpenditures"
            Call mnuExpenditures_Click
            
        Case "mnuCustomersDeposit"
            Call mnuCustomersDeposit_Click

        Case "mnuEmployeesReport"
            Call mnuViewEmployeesRep_Click
            
        Case "mnuCustomersReport"
            Call mnuViewCustomersReport_Click
            
        Case "mnuSuppliersReport"
            Call mnuViewSuppliersAndProduct_Click
            
        Case "mnuExpendituresReport"
            Call mnuViewExpendituresReport_Click
            
        Case "mnuViewSalesReport"
            Call mnuViewSalesReports_Click
            
        Case "mnuTaxReport"
            Call mnuTaxReport_Click

        Case "mnuViewOrdersReports"
            Call mnuViewOrdersReports_Click
            
        Case "mnuPendingOrdersReport"
            Call mnuPendingOrdersReport_Click
            
        Case "mnuOrdersNotPending"
            Call mnuOrdersNotPending_Click
            
        Case "mnuDeliveryReports"
            Call mnuViewDeliveryReports_Click
 
        Case "mnuViewDeliveryTaxReport"
            Call mnuViewDeliveryTaxReport_Click
            
        Case "mnuViewStockReport"
            Call mnuViewStockReport_Click
            
        Case "mnuViewProductsPrice"
            Call mnuViewProductsPrice_Click
            
        Case "mnuActiveProducts"
            Call mnuActiveProducts_Click
            
        Case "mnuInactiveProducts"
            Call mnuInactiveProducts_Click
        
        Case "mnuViewProductsToReorder"
            Call mnuViewProductsToReorder_Click
            
        Case "mnuViewProductsOutOfStock"
            Call mnuViewProductsOutOfStock_Click

        Case "mnuViewProfitLoss"
            Call mnuViewProfitAndLoss_Click

        Case "mnuBankTransaction"
            Call mnuViewBankTransactions_Click
            
        Case "mnuViewUsersTimeRecords"
            Call mnuViewUsersAccessLogs_Click
            
            
            
        Case "mnuCompanyInfo"
            Call mnuViewCompanyInfo_Click
            
        Case "mnuViewEmployees"
            Call mnuViewAllEmployees_Click
            
        Case "mnuViewLeaves"
            Call mnuLeaves_Click
            
        Case "mnuViewSalaries"
            Call mnuSalaries_Click
            
        Case "mnuViewSalaryPeriod"
            Call mnuViewSalaryPeriod_Click
            
        Case "mnuCustomers"
            Call mnuViewAllCustomers_Click
            
        Case "mnuSuppliers"
            Call mnuAllSuppliers_Click
            
        Case "mnuCategories"
            Call mnuViewAllCategories_Click

        Case "mnuProducts"
            Call mnuViewAllProducts_Click
            
        Case "mnuUsers"
            Call mnuViewAllUsers_Click

        Case "mnuBankAndAccount"
            Call mnuBankAndAccount_Click
            
        Case "mnuDeposit"
            Call mnuBankDeposit_Click
        
        Case "mnuWithdrawal"
            Call mnuBankWithdrawal_Click
        
        Case "mnuCharges"
            Call mnuBankCharges_Click
            
        Case "mnuCalculator"
            Call mnuCalculator_Click

        Case "mnuNotepad"
            Call mnunotepad_Click
            
    End Select

End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    Call sub_LOAD_FORMS(Node.Key)
End Sub

Public Sub sub_DISABLE_FEATURES()

    Dim mMenu As Control
    
    For Each mMenu In Me
        If TypeOf mMenu Is Menu Then
            If mMenu.Caption <> "-" Then
                mMenu.Enabled = False
            End If
        End If
    Next
    
    
    mnuFile.Enabled = True
    mnuChangeUserName.Enabled = True
    mnuLogOff.Enabled = True
    mnuLockApplication.Enabled = True
    mnuQuitApplication.Enabled = True
    mnuTansactions.Enabled = True
    mnuReports.Enabled = True
    mnuAdministration.Enabled = True
    mnuTools.Enabled = True
    mnuWindows.Enabled = True
    mnuCloseAllOpenedWindows.Enabled = True

    mnuSalesReports.Enabled = True
    mnuOrdersReports.Enabled = True
    mnuDeliveryReports.Enabled = True
    mnuUsersReports.Enabled = True

    mnuCustomers.Enabled = True
    mnuCategories.Enabled = True
    mnuSuppliers.Enabled = True
    mnuProducts.Enabled = True
    mnuUsers.Enabled = True

    mnuStockReport.Enabled = True
    mnuCalculator.Enabled = True
    mnuNotepad.Enabled = True
    mnuCompanyInfo.Enabled = True
    mnuCustomersReport.Enabled = True
    mnuSuppliersReport.Enabled = True
    mnuExpendituresReport.Enabled = True
    mnuSales.Enabled = True
    mnuOrders.Enabled = True
    mnuBanks.Enabled = True
    mnuEmployees.Enabled = True
    mnuEmployeesRep.Enabled = True
    mnuBankReport.Enabled = True
    mnuProfitAndLoss.Enabled = True
    
End Sub
