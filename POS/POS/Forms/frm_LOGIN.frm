VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_LOGIN 
   BorderStyle     =   0  'None
   Caption         =   "LOGIN"
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin OCX.b8SideTab b8SideTab1 
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4683
      Caption         =   "LOGIN"
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
      Begin OCX.b8Container b8Container1 
         Height          =   2055
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3625
         BorderColor     =   12735512
         BackColor       =   16185592
         Begin lvButton.lvButtons_H cmdQuit 
            Cancel          =   -1  'True
            Height          =   495
            Left            =   3600
            TabIndex        =   3
            Top             =   1440
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Caption         =   "&Quit"
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
            Image           =   "frm_LOGIN.frx":0000
            ImgSize         =   32
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_LOGIN.frx":0590
         End
         Begin lvButton.lvButtons_H cmdLogin 
            Height          =   495
            Left            =   2040
            TabIndex        =   2
            Top             =   1440
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Caption         =   "&Login"
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
            Image           =   "frm_LOGIN.frx":08AA
            ImgSize         =   32
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_LOGIN.frx":2584
         End
         Begin OCX.b8Line b8Line1 
            Height          =   60
            Left            =   120
            TabIndex        =   8
            Top             =   1320
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   106
            BorderColor1    =   12735512
         End
         Begin VB.TextBox txtPassword 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2040
            MaxLength       =   20
            PasswordChar    =   "v"
            TabIndex        =   1
            Text            =   "xx"
            Top             =   720
            Width           =   3045
         End
         Begin VB.TextBox txtUserName 
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
            Left            =   2040
            MaxLength       =   20
            TabIndex        =   0
            Text            =   "Abdel"
            Top             =   240
            Width           =   3045
         End
         Begin VB.Image Image1 
            Height          =   975
            Left            =   120
            Picture         =   "frm_LOGIN.frx":289E
            Stretch         =   -1  'True
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
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
            Left            =   840
            TabIndex        =   7
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "User Name"
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
            Left            =   840
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frm_LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngUserID As Long
Dim lngAdminID As Long

Private Sub cmdQuit_Click()
    Unload Me
    Unload frm_MAIN
End Sub

Private Sub cmdLogin_Click()
On Error GoTo errHandler

    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtUserName, "Kindly enter user name.") Then GoTo EXITPROCEDURE
        
    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtPassword, "Kindly enter password.") Then GoTo EXITPROCEDURE
    
    With cls_USER_Obj
        If .fn_CHECK_USER_LOGIN(Trim(txtUserName), Trim(txtPassword)) Then
                lngCurrentUserID = .UserID
                lngUserID = .UserID
                lngAdminID = .Admin
                strUserName = .UserName
                strPassword = .Password
                strFullName = .FullName
                Call sub_SAVE_USERS_LOGS(1, "Log In")
                Call sub_LOAD_FEATURES(lngUserID)
                Call sub_LOAD_VAT_NHIL
                
                With frm_MAIN
                    If lngAdminID = 1 Then
                        .lblRole.Caption = "Administrator"
                        Else
                            .lblRole.Caption = "Role"
                    End If
                    .lblName.Caption = strUserName
                    .lblTime.Caption = Now
                End With
                Unload Me
            Else
                MsgBox "Invalid user name and password.", vbExclamation, Title
                Call Mdl_FUNCTIONS.fn_HIGHLIGHT_TEXT(txtUserName)
        End If
    End With
    
EXITPROCEDURE:
    Exit Sub
    
errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn
            Call cmdLogin_Click
    End Select

End Sub

Private Sub Form_Load()

    Call Mdl_FUNCTIONS.sub_CENTER_FORM(Me)
    Call frm_MAIN.sub_DISABLE_FEATURES
    
End Sub


Private Sub txtUserName_GotFocus()
    Call Mdl_FUNCTIONS.fn_HIGHLIGHT_TEXT(txtUserName)
End Sub

Private Sub txtPassword_GotFocus()
    Call Mdl_FUNCTIONS.fn_HIGHLIGHT_TEXT(txtPassword)
End Sub


Public Sub sub_SAVE_USERS_LOGS(LogType As Long, strDescription As String)
On Error Resume Next

    With cls_USERS_ACCESS_LOG_Obj
        .AccessLogID = .fn_AUTOGEN
        .UserID = lngCurrentUserID
        .WorkStationName = fn_COMPUTER_NAME
        .LoginDate = Format(Date, "dd/MM/yyyy")
        .LoginTime = Now
        .LoginType = LogType
        .Description = strDescription
        .fn_SAVE_USERS_ACCESS_LOG
    End With


End Sub

Private Sub sub_LOAD_VAT_NHIL()

    Dim rec As New ADODB.Recordset
    
    Set rec = cls_COMPANY_INFO_Obj.fn_LOAD_COMPANY
    If rec.AbsolutePosition <> -1 Then
        lngVAT = Trim(rec!VATRate)
        lngNHIL = Trim(rec!NHILRate)
        strCompanyName = Trim(rec!CompanyName)
        strAddress = Trim(rec!Address)
        strEMail = Trim(rec!EMail)
        strPhoneNo = Trim(rec!PhoneNo)
        strFax = Trim(rec!Fax)
        strLocation = Trim(rec!Location)
        strVatNo = Trim(rec!VATNO)

    End If
    
End Sub


Private Sub sub_LOAD_FEATURES(lngUserID As Long)

    Dim rec As New ADODB.Recordset
    Set rec = cls_USER_Obj.fn_LOAD_FEATURES(lngUserID)
    
    Do While Not rec.EOF
        With frm_MAIN
            Select Case rec!FeatureID
                Case 1
                    .mnuCashSales.Enabled = True
                Case 2
                    .mnuOrdersToSuppliers.Enabled = True
                Case 3
                    .mnuDelivery.Enabled = True
                Case 4
                    .mnuViewSalesReports.Enabled = True
                Case 5
                    .mnuViewOrdersReports.Enabled = True
                Case 6
                    .mnuPendingOrdersReport.Enabled = True
                Case 7
                    .mnuOrdersNotPending.Enabled = True
                Case 8
                    .mnuViewDeliveryReports.Enabled = True
                Case 9
                    .mnuViewStockReport.Enabled = True
                Case 10
                    .mnuViewUsersAccessLogs.Enabled = True
                Case 11
                    .mnuViewAllCustomers.Enabled = True
                Case 12
                    .mnuAllSuppliers.Enabled = True
                Case 13
                    .mnuViewAllCategories.Enabled = True
                Case 14
                    .mnuViewAllProducts.Enabled = True
                Case 15
                    .mnuViewAllUsers.Enabled = True
                Case 16
                    .mnuViewCompanyInfo.Enabled = True
                Case 17
                    .mnuActiveProducts.Enabled = True
                Case 18
                    .mnuInactiveProducts.Enabled = True
                Case 19
                    .mnuViewProductsOutOfStock.Enabled = True
                Case 20
                    .mnuViewProductsToReorder.Enabled = True
                Case 21
                    .mnuViewCustomersReport.Enabled = True
                Case 22
                    .mnuViewSuppliersAndProduct.Enabled = True
                Case 23
                    .mnuExpenditures.Enabled = True
                Case 24
                    .mnuViewExpendituresReport.Enabled = True
                Case 25
                    .mnuAllCreditSales.Enabled = True
                Case 26
                    .mnuOrdersFromCustomers.Enabled = True
                Case 27
                    .mnuCustomersDeposit.Enabled = True
                Case 28
                    .mnuBankAndAccount.Enabled = True
                Case 29
                    .mnuCreditSales.Enabled = True
                Case 30
                    .mnuViewAllEmployees.Enabled = True
                Case 31
                    .mnuLeaves.Enabled = True
                Case 32
                    .mnuSalaries.Enabled = True
                Case 33
                    .mnuViewEmployeesRep.Enabled = True
'                Case 34
'                    .mnuleavesReport.Enabled = True
'                Case 35
'                    .mnuSalariesReport.Enabled = True
                Case 36
                    .mnuTaxReport.Enabled = True
                Case 37
                    .mnuBankDeposit.Enabled = True
                Case 38
                    .mnuBankWithdrawal.Enabled = True
                Case 39
                    .mnuBankCharges.Enabled = True
                Case 40
                    .mnuViewBankTransactions.Enabled = True
                Case 41
                    .mnuViewSalaryPeriod.Enabled = True
                Case 42
                    .mnuViewDeliveryTaxReport.Enabled = True
                Case 43
                    .mnuServerConnection.Enabled = True
                Case 44
                    .mnuViewProductsPrice.Enabled = True
            End Select
        End With
        rec.MoveNext
    Loop
    
    
    
End Sub
