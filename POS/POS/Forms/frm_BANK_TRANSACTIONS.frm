VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_BANK_TRANSACTIONS 
   Caption         =   "BANK TRANSACTIONS"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   12180
   Begin OCX.b8Container b8Container1 
      Height          =   8745
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   15425
      BackColor       =   16185592
      Begin OCX.b8Container b8Container4 
         Height          =   795
         Left            =   90
         TabIndex        =   23
         Top             =   7860
         Width           =   11985
         _ExtentX        =   21140
         _ExtentY        =   1402
         BackColor       =   16185592
         Begin lvButton.lvButtons_H cmdNewDeposit 
            Height          =   525
            Left            =   180
            TabIndex        =   24
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   926
            Caption         =   "New Deposit"
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
            Image           =   "frm_BANK_TRANSACTIONS.frx":0000
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANK_TRANSACTIONS.frx":0515
         End
         Begin lvButton.lvButtons_H cmdNewWithdrawal 
            Height          =   525
            Left            =   1880
            TabIndex        =   25
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   926
            Caption         =   "&New Withdrawal"
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
            Image           =   "frm_BANK_TRANSACTIONS.frx":082F
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANK_TRANSACTIONS.frx":0D44
         End
         Begin lvButton.lvButtons_H cmdNewBankCharges 
            Height          =   525
            Left            =   3580
            TabIndex        =   26
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   926
            Caption         =   "&New Bank Charges"
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
            Image           =   "frm_BANK_TRANSACTIONS.frx":105E
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANK_TRANSACTIONS.frx":1573
         End
         Begin lvButton.lvButtons_H cmdEdit 
            Height          =   525
            Left            =   5280
            TabIndex        =   27
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   926
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
            Image           =   "frm_BANK_TRANSACTIONS.frx":188D
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANK_TRANSACTIONS.frx":1C0A
         End
         Begin lvButton.lvButtons_H cmdPost 
            Height          =   525
            Left            =   8680
            TabIndex        =   28
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   926
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
            Image           =   "frm_BANK_TRANSACTIONS.frx":1F24
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANK_TRANSACTIONS.frx":2184
         End
         Begin lvButton.lvButtons_H cmdDelete 
            Height          =   525
            Left            =   6980
            TabIndex        =   29
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   926
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
            Image           =   "frm_BANK_TRANSACTIONS.frx":249E
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANK_TRANSACTIONS.frx":29B1
         End
         Begin lvButton.lvButtons_H cmdClose 
            Height          =   525
            Left            =   10380
            TabIndex        =   30
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   926
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
            Image           =   "frm_BANK_TRANSACTIONS.frx":2CCB
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANK_TRANSACTIONS.frx":325B
         End
      End
      Begin OCX.b8Container b8Container3 
         Height          =   2055
         Left            =   90
         TabIndex        =   3
         Top             =   90
         Width           =   11985
         _ExtentX        =   21140
         _ExtentY        =   3625
         BackColor       =   16185592
         Begin MSComCtl2.DTPicker DTPToDate 
            Height          =   345
            Left            =   4440
            TabIndex        =   8
            Top             =   1050
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   609
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
            Format          =   58916865
            CurrentDate     =   39275
         End
         Begin MSComCtl2.DTPicker DTPFromDate 
            Height          =   345
            Left            =   4440
            TabIndex        =   7
            Top             =   570
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   609
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
            Format          =   58916865
            CurrentDate     =   39275
         End
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
            ItemData        =   "frm_BANK_TRANSACTIONS.frx":3575
            Left            =   4440
            List            =   "frm_BANK_TRANSACTIONS.frx":3577
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   150
            Width           =   2955
         End
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
            Height          =   1740
            ItemData        =   "frm_BANK_TRANSACTIONS.frx":3579
            Left            =   210
            List            =   "frm_BANK_TRANSACTIONS.frx":357B
            Sorted          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   4
            Text            =   "cboBank"
            Top             =   150
            Width           =   2835
         End
         Begin lvButton.lvButtons_H cmdLoad 
            Height          =   435
            Left            =   4440
            TabIndex        =   11
            Top             =   1500
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   767
            Caption         =   "È"
            CapAlign        =   2
            BackStyle       =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Wingdings 3"
               Size            =   15.75
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_BANK_TRANSACTIONS.frx":357D
         End
         Begin VB.Label Label2 
            BackColor       =   &H00F6F8F8&
            Caption         =   "To Date"
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
            Left            =   3150
            TabIndex        =   10
            Top             =   1110
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackColor       =   &H00F6F8F8&
            Caption         =   "From Date"
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
            Left            =   3150
            TabIndex        =   9
            Top             =   630
            Width           =   1215
         End
         Begin VB.Label Label7 
            BackColor       =   &H00F6F8F8&
            Caption         =   "Account No"
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
            Left            =   3150
            TabIndex        =   6
            Top             =   180
            Width           =   1215
         End
      End
      Begin OCX.b8Container b8Container2 
         Height          =   5715
         Left            =   90
         TabIndex        =   1
         Top             =   2160
         Width           =   11985
         _ExtentX        =   21140
         _ExtentY        =   10081
         BackColor       =   16185592
         Begin MSComctlLib.ListView lvw 
            Height          =   4095
            Left            =   210
            TabIndex        =   2
            Top             =   60
            Width           =   11625
            _ExtentX        =   20505
            _ExtentY        =   7223
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "TransactionID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "TransactionNo"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "TransactionType"
               Object.Width           =   2470
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Status"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Description"
               Object.Width           =   6175
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Amount"
               Object.Width           =   4057
            EndProperty
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            ForeColor       =   &H80000008&
            Height          =   1515
            Left            =   7740
            TabIndex        =   12
            Top             =   4140
            Width           =   4185
            Begin VB.TextBox txtPostedBalance 
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
            Begin VB.TextBox txtActualBalance 
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
               TabIndex        =   19
               Top             =   900
               Width           =   2265
            End
            Begin VB.TextBox txtTotalCharges 
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
               TabIndex        =   17
               Top             =   600
               Width           =   2265
            End
            Begin VB.TextBox txtTotalWithdrawal 
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
               TabIndex        =   15
               Top             =   300
               Width           =   2265
            End
            Begin VB.TextBox txtTotalDeposit 
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
               TabIndex        =   13
               Top             =   0
               Width           =   2265
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Posted balance"
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
               TabIndex        =   22
               Top             =   1230
               Width           =   1515
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Actual balance"
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
               TabIndex        =   20
               Top             =   930
               Width           =   1515
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Charges"
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
               TabIndex        =   18
               Top             =   630
               Width           =   1425
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Withdrawal"
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
               TabIndex        =   16
               Top             =   330
               Width           =   1545
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Deposit"
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
               TabIndex        =   14
               Top             =   30
               Width           =   1425
            End
         End
      End
   End
End
Attribute VB_Name = "frm_BANK_TRANSACTIONS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngID As Long
Dim lngAccountID As Long

Private Sub cboAccount_Click()
    If cboAccount.ListIndex = -1 Then Exit Sub
    lngAccountID = cboAccount.ItemData(cboAccount.ListIndex)
End Sub

Private Sub cboBank_Click()
    If cboBank.ListIndex = -1 Then Exit Sub
    lngID = cboBank.ItemData(cboBank.ListIndex)
    Call loadBankAccounts(lngID)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If lvw.ListItems.Count = 0 Then
        MsgBox "Please select the transaction to delete", vbExclamation, Title
        Exit Sub
    End If
    
    If Trim(lvw.SelectedItem.ListSubItems(4).Text) = "Posted" Then
        MsgBox "Can not delete a posted transaction", vbInformation, Title
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to post this transaction?", vbYesNo + vbQuestion, Title) = vbYes Then
        Call cls_BANK_TRANSACTION_Obj.fn_DELETE_BANK_RANSACTIONS(lvw.SelectedItem.Text)
        Call cmdLoad_Click
    End If

End Sub

Private Sub cmdEdit_Click()

    If lvw.ListItems.Count = 0 Then
        MsgBox "Please select the transaction to edit", vbExclamation, Title
        Exit Sub
    End If
    
    If Trim(lvw.SelectedItem.ListSubItems(4).Text) = "Posted" Then
        MsgBox "Can not edit a posted transaction", vbInformation, Title
        Exit Sub
    End If
    
    
End Sub

Public Sub cmdLoad_Click()
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_COMBO(cboBank, "Please select the bank name") = True Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_COMBO(cboAccount, "Please select the account no") = True Then Exit Sub

    If DTPFromDate.Value > Date Then
        MsgBox "The starting date should not be more than today", vbExclamation, Title
        DTPFromDate.SetFocus
        Exit Sub
    End If
    
    If DTPFromDate.Value > DTPToDate.Value Then
        MsgBox "The starting date should not be more than the ending date", vbExclamation, Title
        DTPFromDate.SetFocus
        Exit Sub
    End If
    
    Call sub_LOAD_BANK_TRANSACTIONS(lngAccountID, DTPFromDate.Value, DTPToDate.Value)
    
End Sub

Private Sub sub_LOAD_BANK_TRANSACTIONS(Optional lngAccountID As Long, Optional fromDate As Date, Optional toDate As Date)
    Dim ctr As Long
    
    Dim rec As New ADODB.Recordset
    Set rec = cls_BANK_TRANSACTION_Obj.fn_LOAD_BANK_TRANSACTIONS(lngAccountID, fromDate, toDate)
    lvw.ListItems.Clear
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            Set lstItem = lvw.ListItems.Add(, , rec!TransactionID)
                lstItem.ListSubItems.Add , , Trim(rec!TransactionDate)
                lstItem.ListSubItems.Add , , Trim(rec!TransactionNo)
                lstItem.ListSubItems.Add , , Trim(rec!TransactionType)
                If rec!Posted = 0 Then
                    lstItem.ListSubItems.Add , , "Pending"
                    Else
                        lstItem.ListSubItems.Add , , "Posted"
                End If
                lstItem.ListSubItems.Add , , Trim(rec!Description)
                lstItem.ListSubItems.Add , , Trim(rec!Amount)
        rec.MoveNext
        Loop
    End If
    
    txtTotalDeposit.Text = ""
    txtTotalWithdrawal.Text = ""
    txtTotalCharges.Text = ""
    txtPostedBalance.Text = ""
    
    For ctr = 1 To lvw.ListItems.Count
        If lvw.ListItems(ctr).ListSubItems(3).Text = "Deposit" Then
            txtTotalDeposit.Text = Val(txtTotalDeposit.Text) + Val(lvw.ListItems(ctr).ListSubItems(6).Text)
        End If
        If lvw.ListItems(ctr).ListSubItems(3).Text = "Withdrawal" Then
            txtTotalWithdrawal.Text = Val(txtTotalWithdrawal.Text) + Val(lvw.ListItems(ctr).ListSubItems(6).Text)
        End If
        If lvw.ListItems(ctr).ListSubItems(3).Text = "Bank Charges" Then
            txtTotalCharges.Text = Val(txtTotalCharges.Text) + Val(lvw.ListItems(ctr).ListSubItems(6).Text)
        End If
        If lvw.ListItems(ctr).ListSubItems(3).Text = "Deposit" And lvw.ListItems(ctr).ListSubItems(4).Text = "Posted" Then
            txtPostedBalance.Text = Val(txtPostedBalance.Text) + Val(lvw.ListItems(ctr).ListSubItems(6).Text)
        End If
    Next
    
    txtActualBalance.Text = Val(txtTotalDeposit.Text) - Val(Val(txtTotalWithdrawal.Text) + Val(txtTotalCharges.Text))
    
End Sub

Private Sub cmdNewBankCharges_Click()
    If frm_MAIN.mnuBankCharges.Enabled = False Then Exit Sub
    blnNewBankCharges = True
    With frm_BANK_CHARGES
        .Show 1
    End With
End Sub

Private Sub cmdNewDeposit_Click()
    If frm_MAIN.mnuBankDeposit.Enabled = False Then Exit Sub
    blnNewDeposit = True
    With frm_DEPOSIT
        .Show 1
    End With
End Sub

Private Sub cmdNewWithdrawal_Click()
    If frm_MAIN.mnuBankWithdrawal.Enabled = False Then Exit Sub
    blnNewWithdrawal = True
    With frm_WITHDRAWAL
        .Show 1
    End With
End Sub

Private Sub cmdPost_Click()
    If lvw.ListItems.Count = 0 Then
        MsgBox "Please select the transaction to post", vbExclamation, Title
        Exit Sub
    End If
    
    If Trim(lvw.SelectedItem.ListSubItems(4).Text) = "Posted" Then
        MsgBox "Can not post a posted transaction", vbInformation, Title
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to post this transaction?", vbYesNo + vbQuestion, Title) = vbYes Then
        Call cls_BANK_TRANSACTION_Obj.fn_POST_BANK_RANSACTIONS(lvw.SelectedItem.Text)
        Call cmdLoad_Click
    End If
End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_FORM_SIZE(Me)
    Move 0, 0
    Call subLoadBank(cboBank)
    DTPFromDate.Value = Date
    DTPToDate.Value = Date
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
