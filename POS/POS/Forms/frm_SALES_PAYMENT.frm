VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_SALES_PAYMENT 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PAYMENT"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin OCX.b8Container b8Container2 
      Height          =   705
      Left            =   0
      TabIndex        =   10
      Top             =   3360
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1244
      BackColor       =   16185592
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1920
         TabIndex        =   11
         Top             =   120
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   873
         Caption         =   "&Post"
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
         Image           =   "frm_SALES_PAYMENT.frx":0000
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SALES_PAYMENT.frx":0260
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   3300
         TabIndex        =   12
         Top             =   120
         Width           =   1275
         _ExtentX        =   2249
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
         Image           =   "frm_SALES_PAYMENT.frx":057A
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SALES_PAYMENT.frx":0B0A
      End
   End
   Begin OCX.b8Container b8Container1 
      Height          =   3345
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5900
      BackColor       =   16185592
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3090
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4695
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
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   0
            Width           =   3045
         End
         Begin VB.TextBox txtDiscount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1485
            TabIndex        =   0
            Top             =   1320
            Width           =   3000
         End
         Begin VB.TextBox TxtAmountPaid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1485
            TabIndex        =   3
            Top             =   2220
            Width           =   3000
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
            Left            =   60
            TabIndex        =   18
            Top             =   15
            Width           =   1545
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Credit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   60
            TabIndex        =   17
            Top             =   420
            Width           =   1245
         End
         Begin VB.Label lblAvailableAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1485
            TabIndex        =   16
            Top             =   420
            Width           =   3015
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
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
            Left            =   60
            TabIndex        =   14
            Top             =   2670
            Width           =   1185
         End
         Begin VB.Label LblChange 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   1485
            TabIndex        =   13
            Top             =   2670
            Width           =   3000
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Gross Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   30
            TabIndex        =   9
            Top             =   870
            Width           =   1305
         End
         Begin VB.Label LblGrossAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   1485
            TabIndex        =   8
            Top             =   870
            Width           =   3000
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Net Amount"
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
            Left            =   60
            TabIndex        =   7
            Top             =   1770
            Width           =   1245
         End
         Begin VB.Label LblNetAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   1485
            TabIndex        =   6
            Top             =   1770
            Width           =   3000
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount Paid"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   5
            Top             =   2220
            Width           =   1245
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Discount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   30
            TabIndex        =   4
            Top             =   1320
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frm_SALES_PAYMENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strCompanyInfo As String
Dim lngCustomerID As Long
Private Sub subLoadCustomersDetails(lngID As Long)

    Dim rec As New ADODB.Recordset
    Set rec = cls_CUSTOMER_Obj.fn_LOAD_CUSTOMERS(lngID)

    If rec.AbsolutePosition <> -1 Then
        lblAvailableAmount.Caption = Val(rec!Amount)
        TxtAmountPaid.Text = Val(rec!Amount)
    End If
    
    If Trim(lblAvailableAmount.Caption) = 0 Or Trim(lblAvailableAmount.Caption) = "" Then
        TxtAmountPaid.Enabled = True
        
        Else
            TxtAmountPaid.Enabled = False
            
    End If
    
End Sub

Private Sub cboCustomer_Click()
    If cboCustomer.ListIndex = -1 Then
        lngCustomerID = 0
        Else
            lngCustomerID = cboCustomer.ItemData(cboCustomer.ListIndex)
    End If
    Call subLoadCustomersDetails(lngCustomerID)
End Sub

Private Sub cmdCancel_Click()
    blnCancelPayment = True
    Unload Me
End Sub


Private Sub cmdSave_Click()

    If Val(LblChange.Caption) < 0 Then
        MsgBox "The Amount paid Is Invalid. Please make the sale credit", vbExclamation, Title
        TxtAmountPaid.SetFocus
        Exit Sub
    End If

    If Val(TxtAmountPaid.Text) <= 0 Then
        MsgBox "The Amount paid Is Invalid", vbExclamation, Title
        TxtAmountPaid.SetFocus
    Else
            With frm_SALES
                .dblGrossAmount = Val(LblGrossAmount.Caption)
                .dblDiscount = Val(txtDiscount.Text)
                .dblNetAmount = Val(LblNetAmount.Caption)
                .dblAmountPaid = Val(TxtAmountPaid.Text)
                .dblBalance = Val(LblChange.Caption)
            End With
            
            Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_CENTER_FORM(Me)
    Call sub_LOAD_CUSTOMERS
    With frm_SALES
        .dblGrossAmount = 0
        .dblDiscount = 0
        .dblNetAmount = 0
        .dblAmountPaid = 0
        .dblBalance = 0
    End With
    
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSave, cmdSave.hwnd, "Post sales details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
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

Private Sub TxtAmountPaid_Change()
    LblChange.Caption = Val(TxtAmountPaid.Text) - Val(LblNetAmount.Caption)
End Sub

Private Sub TxtAmountPaid_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtDiscount_Change()
    LblNetAmount.Caption = Val(LblGrossAmount.Caption) - Val(txtDiscount.Text)
'    LblChange.Caption = Val(TxtAmountPaid.Text) + Val(txtDiscount.Text)
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub
