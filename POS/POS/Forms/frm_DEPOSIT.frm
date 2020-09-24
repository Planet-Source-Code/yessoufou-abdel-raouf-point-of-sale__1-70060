VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_DEPOSIT 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DEPOSIT"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   Begin OCX.b8Container b8Container2 
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   5160
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   1296
      BackColor       =   16185592
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   7935
         TabIndex        =   11
         Top             =   120
         Width           =   1365
         _ExtentX        =   2408
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
         Image           =   "frm_DEPOSIT.frx":0000
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_DEPOSIT.frx":056F
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   6420
         TabIndex        =   10
         Top             =   120
         Width           =   1365
         _ExtentX        =   2408
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
         Image           =   "frm_DEPOSIT.frx":0889
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_DEPOSIT.frx":0AE9
      End
   End
   Begin OCX.b8Container b8Container1 
      Height          =   5145
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   9075
      BackColor       =   16185592
      Begin VB.Frame Frame2 
         BackColor       =   &H00F6F8F8&
         Height          =   4965
         Left            =   90
         TabIndex        =   26
         Top             =   90
         Width           =   3015
         Begin VB.ComboBox cboBank 
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
            Height          =   4665
            ItemData        =   "frm_DEPOSIT.frx":0E03
            Left            =   90
            List            =   "frm_DEPOSIT.frx":0E05
            Sorted          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   27
            Text            =   "cboBank"
            Top             =   180
            Width           =   2805
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00F6F8F8&
         Height          =   3225
         Left            =   3150
         TabIndex        =   18
         Top             =   90
         Width           =   6165
         Begin VB.ComboBox cboAccount 
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
            ItemData        =   "frm_DEPOSIT.frx":0E07
            Left            =   1500
            List            =   "frm_DEPOSIT.frx":0E09
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   4545
         End
         Begin VB.TextBox TxtAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1500
            MaxLength       =   13
            TabIndex        =   2
            Top             =   990
            Width           =   4515
         End
         Begin VB.TextBox TxtDepositedBy 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1500
            MaxLength       =   50
            TabIndex        =   3
            Top             =   1350
            Width           =   4515
         End
         Begin VB.TextBox TxtDescription 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   1500
            MaxLength       =   100
            TabIndex        =   4
            Top             =   1710
            Width           =   4515
         End
         Begin VB.ComboBox CboMode 
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
            ItemData        =   "frm_DEPOSIT.frx":0E0B
            Left            =   1500
            List            =   "frm_DEPOSIT.frx":0E15
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2430
            Width           =   4545
         End
         Begin VB.CheckBox ChkPostDeposit 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1500
            TabIndex        =   6
            Top             =   2760
            Width           =   225
         End
         Begin MSComCtl2.DTPicker DtpTransactionDate 
            Height          =   315
            Left            =   1500
            TabIndex        =   1
            Top             =   600
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
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
            CustomFormat    =   "dd MMM yyyy"
            Format          =   58916867
            CurrentDate     =   37962
         End
         Begin VB.Label Label7 
            BackColor       =   &H00F6F8F8&
            Caption         =   "Account"
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
            Left            =   150
            TabIndex        =   25
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label LblDate 
            BackColor       =   &H00F6F8F8&
            Caption         =   "Date Deposited"
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
            TabIndex        =   24
            Top             =   630
            Width           =   1335
         End
         Begin VB.Label Label5 
            BackColor       =   &H00F6F8F8&
            Caption         =   "Amount"
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
            TabIndex        =   23
            Top             =   1020
            Width           =   1095
         End
         Begin VB.Label LblBy 
            AutoSize        =   -1  'True
            BackColor       =   &H00F6F8F8&
            Caption         =   "Deposited By"
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
            Left            =   150
            TabIndex        =   22
            Top             =   1410
            Width           =   1140
         End
         Begin VB.Label LblMode 
            BackColor       =   &H00F6F8F8&
            Caption         =   "Deposit Mode"
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
            Left            =   180
            TabIndex        =   21
            Top             =   2460
            Width           =   1215
         End
         Begin VB.Label Label8 
            BackColor       =   &H00F6F8F8&
            Caption         =   "Remarks"
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
            TabIndex        =   20
            Top             =   1740
            Width           =   1065
         End
         Begin VB.Label Label4 
            BackColor       =   &H00F6F8F8&
            Caption         =   "Post Deposit"
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
            Left            =   180
            TabIndex        =   19
            Top             =   2790
            Width           =   1125
         End
      End
      Begin VB.Frame FraCheque 
         BackColor       =   &H00F6F8F8&
         Caption         =   "Cheque Details"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Left            =   3150
         TabIndex        =   14
         Top             =   3450
         Width           =   6165
         Begin VB.ComboBox cboBanks 
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
            TabIndex        =   7
            Top             =   300
            Width           =   4545
         End
         Begin VB.TextBox TxtChequeNo 
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
            MaxLength       =   15
            TabIndex        =   9
            Top             =   1020
            Width           =   4515
         End
         Begin MSComCtl2.DTPicker DtpChequeDate 
            Height          =   315
            Left            =   1500
            TabIndex        =   8
            Top             =   660
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   556
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
            CustomFormat    =   "dd MMM yyyy"
            Format          =   58916867
            CurrentDate     =   37962
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank"
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
            Left            =   150
            TabIndex        =   17
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cheque Date"
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
            Top             =   690
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Cheque No."
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
            Left            =   150
            TabIndex        =   15
            Top             =   1050
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "frm_DEPOSIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngID As Long
Dim lngAccountID As Long
Dim blnCash As Boolean
Dim blnCheque As Boolean
Private Sub cboAccount_Click()
    If cboAccount.ListIndex = -1 Then Exit Sub
    lngAccountID = cboAccount.ItemData(cboAccount.ListIndex)
    
End Sub

Private Sub cboBank_Click()
    If cboBank.ListIndex = -1 Then Exit Sub
    lngID = cboBank.ItemData(cboBank.ListIndex)
    Call loadBankAccounts(lngID)
End Sub

Private Sub CboMode_Click()
    Select Case CboMode.List(CboMode.ListIndex)
        Case "Cash"
            blnCash = True
            blnCheque = False
            FraCheque.Enabled = False
        Case "Cheque"
            blnCheque = True
            blnCash = False
            FraCheque.Enabled = True
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_COMBO(cboAccount, "Please select the account.") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(TxtAmount, "Please enter the amount.") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_COMBO(CboMode, "Please select deposit mode.") Then Exit Sub
'    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtLastName, "Please enter last name.") Then Exit Sub

    With cls_BANK_TRANSACTION_Obj
        .TransactionDate = DtpTransactionDate.Value
        .TransactionType = "Deposit"
        .Description = Trim(TxtDescription.Text)
        .AccountID = lngAccountID
        .Amount = Val(TxtAmount.Text)
        .TransactedBy = Trim(TxtDepositedBy.Text)
        If ChkPostDeposit.Value = 1 Then
            .Posted = 1
            Else
                .Posted = 0
        End If
        .ChequeID = 0
        .UserID = lngCurrentUserID
        Call .fn_SAVE_BANK_TRANSACTIONS
    
    End With
    
    With cls_CHEQUES_Obj
        If blnCash = False And blnCheque = True Then
            If Mdl_FUNCTIONS.fn_CHECK_EMPTY_COMBO(cboBank, "Please select the bank.") Then Exit Sub
            If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(TxtChequeNo, "Please enter the cheque no.") Then Exit Sub
            .ChequeDate = DtpChequeDate.Value
            .BankID = cboBanks.ItemData(cboBanks.ListIndex)
            .ChequeNo = Trim(TxtChequeNo.Text)
            .Amount = Trim(TxtAmount.Text)
            .Status = 0
            .Location = 0
            .UserID = lngCurrentUserID
            Call .fn_SAVE_CHEQUES
        End If
    End With
    
    MsgBox "Transaction Saved Successfully", vbInformation, Title
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    If blnNewDeposit = True Then
        Call frm_BANK_TRANSACTIONS.cmdLoad_Click
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_CENTER_FORM(Me)
    DtpTransactionDate.Value = Date
    DtpChequeDate.Value = Date
    Call subLoadBank(cboBank)
    Call subLoadBank(cboBanks)
End Sub

Private Sub subLoadBank(cbo As ComboBox)

    Dim rec As New ADODB.Recordset
    Set rec = cls_BANK_Obj.fn_LOAD_BANKS(0)
    
    cbo.Clear
    
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            cbo.AddItem rec!BankName
            cbo.ItemData(cbo.NewIndex) = rec!BankID
            rec.MoveNext
        Loop
    End If
    
End Sub

Private Sub loadBankAccounts(lngID As Long)

    Dim rec As New ADODB.Recordset
    
    Set rec = cls_BANK_Obj.fn_LOAD_BANKS_ACCOUNTS(lngID)
    
    cboAccount.Clear
    
    If rec.AbsolutePosition = -1 Then Exit Sub
    Do While Not rec.EOF
        cboAccount.AddItem rec!AccountName
        cboAccount.ItemData(cboAccount.NewIndex) = rec!AccountID
        rec.MoveNext
    Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)
    blnNewDeposit = False
End Sub

Private Sub TxtAmount_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub TxtDepositedBy_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_ALPHABET_ONLY(KeyAscii)
End Sub
