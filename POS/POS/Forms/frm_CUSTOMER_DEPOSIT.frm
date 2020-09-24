VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_CUSTOMER_DEPOSIT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CUSTOMERS DEPOSIT"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   6495
   Begin VB.Frame fra 
      BackColor       =   &H00F6F8F8&
      Height          =   6795
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   6495
      Begin TabDlg.SSTab SSTab1 
         Height          =   6525
         Left            =   90
         TabIndex        =   1
         Top             =   150
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   11509
         _Version        =   393216
         TabHeight       =   520
         BackColor       =   16777215
         TabCaption(0)   =   "Cash"
         TabPicture(0)   =   "frm_CUSTOMER_DEPOSIT.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "b8Container1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Cheque"
         TabPicture(1)   =   "frm_CUSTOMER_DEPOSIT.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "b8Container2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Bank Deposit"
         TabPicture(2)   =   "frm_CUSTOMER_DEPOSIT.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "b8Container3"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin OCX.b8Container b8Container3 
            Height          =   6165
            Left            =   -75000
            TabIndex        =   39
            Top             =   360
            Width           =   6315
            _ExtentX        =   11139
            _ExtentY        =   10874
            BackColor       =   16185592
            Begin VB.Frame Frame4 
               BackColor       =   &H00F6F8F8&
               Height          =   645
               Left            =   150
               TabIndex        =   64
               Top             =   5400
               Width           =   5985
               Begin lvButton.lvButtons_H cmdBankDepositCancel 
                  Height          =   405
                  Left            =   3180
                  TabIndex        =   65
                  Top             =   180
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   714
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
                  ImgSize         =   32
                  cBack           =   -2147483633
                  mPointer        =   99
                  mIcon           =   "frm_CUSTOMER_DEPOSIT.frx":0054
               End
               Begin lvButton.lvButtons_H cmdBankDepositSave 
                  Height          =   405
                  Left            =   1650
                  TabIndex        =   66
                  Top             =   180
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   714
                  Caption         =   "&Credit"
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
                  ImgSize         =   32
                  cBack           =   -2147483633
                  mPointer        =   99
                  mIcon           =   "frm_CUSTOMER_DEPOSIT.frx":036E
               End
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00F6F8F8&
               Height          =   5325
               Left            =   150
               TabIndex        =   40
               Top             =   90
               Width           =   6105
               Begin VB.TextBox txtBankDepositCustomer 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1530
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   69
                  Text            =   " "
                  Top             =   780
                  Width           =   2985
               End
               Begin VB.ComboBox cboBankDepositBankName 
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
                  Left            =   1530
                  Style           =   2  'Dropdown List
                  TabIndex        =   50
                  Top             =   1590
                  Width           =   2985
               End
               Begin VB.TextBox txtBankDepositDepositedBy 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1530
                  MaxLength       =   100
                  TabIndex        =   49
                  Text            =   " "
                  Top             =   4830
                  Width           =   2985
               End
               Begin VB.TextBox txtBankDepositChequeNo 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1530
                  MaxLength       =   20
                  TabIndex        =   48
                  Text            =   " "
                  Top             =   1185
                  Width           =   2985
               End
               Begin VB.TextBox txtBankDepositAmount 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1530
                  MaxLength       =   20
                  TabIndex        =   47
                  Text            =   " "
                  Top             =   1980
                  Width           =   2985
               End
               Begin VB.TextBox txtBankDeposited 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1530
                  MaxLength       =   100
                  TabIndex        =   46
                  Text            =   " "
                  Top             =   2370
                  Width           =   2985
               End
               Begin VB.TextBox txtBankDepositBranchCode 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1530
                  MaxLength       =   100
                  TabIndex        =   45
                  Text            =   " "
                  Top             =   2760
                  Width           =   2985
               End
               Begin VB.TextBox txtBankDepositTransactionNo 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1530
                  MaxLength       =   100
                  TabIndex        =   44
                  Text            =   " "
                  Top             =   3150
                  Width           =   2985
               End
               Begin VB.OptionButton optBankDepositCash 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00F6F8F8&
                  Caption         =   "Cash"
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
                  Height          =   285
                  Left            =   180
                  TabIndex        =   43
                  Top             =   270
                  Width           =   1365
               End
               Begin VB.OptionButton optBankDepositCheque 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00F6F8F8&
                  Caption         =   "Cheque"
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
                  Left            =   2370
                  TabIndex        =   42
                  Top             =   270
                  Width           =   1425
               End
               Begin VB.TextBox txtBankDepositAccountNo 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1530
                  MaxLength       =   20
                  TabIndex        =   41
                  Text            =   " "
                  Top             =   3540
                  Width           =   2985
               End
               Begin MSComCtl2.DTPicker DTPBankDepositDateReceived 
                  Height          =   345
                  Left            =   1530
                  TabIndex        =   51
                  Top             =   4380
                  Width           =   3015
                  _ExtentX        =   5318
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
                  CustomFormat    =   "dd/MM/yyyy"
                  Format          =   58916867
                  CurrentDate     =   39257
               End
               Begin MSComCtl2.DTPicker DTPBankDepositDatePaid 
                  Height          =   345
                  Left            =   1530
                  TabIndex        =   52
                  Top             =   3930
                  Width           =   2985
                  _ExtentX        =   5265
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
                  CustomFormat    =   "dd/MM/yyyy"
                  Format          =   58916867
                  CurrentDate     =   39257
               End
               Begin lvButton.lvButtons_H cmdBankDepositCustomers 
                  Height          =   375
                  Left            =   4560
                  TabIndex        =   72
                  Top             =   750
                  Width           =   1515
                  _ExtentX        =   2672
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
                  Image           =   "frm_CUSTOMER_DEPOSIT.frx":0688
                  ImgSize         =   24
                  cBack           =   -2147483633
                  mPointer        =   99
                  mIcon           =   "frm_CUSTOMER_DEPOSIT.frx":0B9D
               End
               Begin VB.Label Label15 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Deposited By"
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
                  TabIndex        =   63
                  Top             =   4830
                  Width           =   1485
               End
               Begin VB.Label Label16 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Date Paid"
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
                  TabIndex        =   62
                  Top             =   3945
                  Width           =   1215
               End
               Begin VB.Label Label17 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Date Received"
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
                  TabIndex        =   61
                  Top             =   4395
                  Width           =   1335
               End
               Begin VB.Label Label19 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Account No"
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
                  TabIndex        =   60
                  Top             =   3540
                  Width           =   1485
               End
               Begin VB.Label Label20 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cheque No"
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
                  TabIndex        =   59
                  Top             =   1185
                  Width           =   1485
               End
               Begin VB.Label Label21 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Bank Name"
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
                  TabIndex        =   58
                  Top             =   1575
                  Width           =   1215
               End
               Begin VB.Label Label22 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Customers"
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
                  TabIndex        =   57
                  Top             =   780
                  Width           =   1215
               End
               Begin VB.Label Label24 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Amount"
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
                  TabIndex        =   56
                  Top             =   1980
                  Width           =   1485
               End
               Begin VB.Label Label25 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Bank Deposited"
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
                  TabIndex        =   55
                  Top             =   2370
                  Width           =   1395
               End
               Begin VB.Label Label26 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Branch Code"
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
                  TabIndex        =   54
                  Top             =   2760
                  Width           =   1395
               End
               Begin VB.Label Label27 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Transaction No"
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
                  TabIndex        =   53
                  Top             =   3150
                  Width           =   1395
               End
            End
         End
         Begin OCX.b8Container b8Container2 
            Height          =   6525
            Left            =   -75000
            TabIndex        =   14
            Top             =   360
            Width           =   6285
            _ExtentX        =   11086
            _ExtentY        =   11509
            BackColor       =   16185592
            Begin VB.Frame Frame2 
               BackColor       =   &H00F6F8F8&
               Height          =   4245
               Left            =   150
               TabIndex        =   19
               Top             =   90
               Width           =   6015
               Begin VB.TextBox txtChequeCustomer 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1410
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   68
                  Text            =   " "
                  Top             =   270
                  Width           =   2985
               End
               Begin VB.TextBox txtChequeChequeNo 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1410
                  MaxLength       =   20
                  TabIndex        =   26
                  Text            =   " "
                  Top             =   705
                  Width           =   2985
               End
               Begin VB.TextBox txtChequePaidBy 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1410
                  MaxLength       =   100
                  TabIndex        =   25
                  Text            =   " "
                  Top             =   3120
                  Width           =   2985
               End
               Begin VB.ComboBox cboChequeBanks 
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
                  Left            =   1410
                  Style           =   2  'Dropdown List
                  TabIndex        =   24
                  Top             =   1110
                  Width           =   2985
               End
               Begin VB.ComboBox cboChequeChequeType 
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
                  ItemData        =   "frm_CUSTOMER_DEPOSIT.frx":0EB7
                  Left            =   1470
                  List            =   "frm_CUSTOMER_DEPOSIT.frx":0EC1
                  Style           =   2  'Dropdown List
                  TabIndex        =   23
                  Top             =   4500
                  Width           =   2985
               End
               Begin VB.ComboBox cboChequeStatus 
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
                  ItemData        =   "frm_CUSTOMER_DEPOSIT.frx":0ED2
                  Left            =   1410
                  List            =   "frm_CUSTOMER_DEPOSIT.frx":0EDC
                  Style           =   2  'Dropdown List
                  TabIndex        =   22
                  Top             =   3495
                  Width           =   2985
               End
               Begin VB.TextBox txtChequeAmount 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1410
                  MaxLength       =   20
                  TabIndex        =   21
                  Text            =   " "
                  Top             =   1890
                  Width           =   2985
               End
               Begin VB.TextBox txtChequeAccountNo 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1410
                  MaxLength       =   20
                  TabIndex        =   20
                  Text            =   " "
                  Top             =   1500
                  Width           =   2985
               End
               Begin MSComCtl2.DTPicker DTPChequeDateReceived 
                  Height          =   345
                  Left            =   1410
                  TabIndex        =   27
                  Top             =   2700
                  Width           =   3015
                  _ExtentX        =   5318
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
                  CustomFormat    =   "dd/MM/yyyy"
                  Format          =   58916867
                  CurrentDate     =   39257
               End
               Begin MSComCtl2.DTPicker DTPChequeChequeDate 
                  Height          =   345
                  Left            =   1410
                  TabIndex        =   28
                  Top             =   2310
                  Width           =   2985
                  _ExtentX        =   5265
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
                  CustomFormat    =   "dd/MM/yyyy"
                  Format          =   58916867
                  CurrentDate     =   39257
               End
               Begin lvButton.lvButtons_H cmdChequeCustomers 
                  Height          =   375
                  Left            =   4440
                  TabIndex        =   71
                  Top             =   240
                  Width           =   1515
                  _ExtentX        =   2672
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
                  Image           =   "frm_CUSTOMER_DEPOSIT.frx":0EF6
                  ImgSize         =   24
                  cBack           =   -2147483633
                  mPointer        =   99
                  mIcon           =   "frm_CUSTOMER_DEPOSIT.frx":140B
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Customers"
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
                  TabIndex        =   38
                  Top             =   300
                  Width           =   1215
               End
               Begin VB.Label Label2 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Bank Name"
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
                  TabIndex        =   37
                  Top             =   1095
                  Width           =   1215
               End
               Begin VB.Label Label3 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cheque No"
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
                  TabIndex        =   36
                  Top             =   705
                  Width           =   1485
               End
               Begin VB.Label Label4 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Account No"
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
                  TabIndex        =   35
                  Top             =   1500
                  Width           =   1485
               End
               Begin VB.Label Label7 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cheque Type"
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
                  TabIndex        =   34
                  Top             =   4500
                  Width           =   1215
               End
               Begin VB.Label Label9 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Date Received"
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
                  TabIndex        =   33
                  Top             =   2715
                  Width           =   1335
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cheque Date"
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
                  TabIndex        =   32
                  Top             =   2325
                  Width           =   1215
               End
               Begin VB.Label Label11 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Paid By"
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
                  Top             =   3120
                  Width           =   1485
               End
               Begin VB.Label Label13 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Status"
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
                  TabIndex        =   30
                  Top             =   3480
                  Width           =   1215
               End
               Begin VB.Label Label23 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Amount"
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
                  TabIndex        =   29
                  Top             =   1890
                  Width           =   1485
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00F6F8F8&
               Height          =   645
               Left            =   150
               TabIndex        =   15
               Top             =   4320
               Width           =   6015
               Begin lvButton.lvButtons_H cmdChequeCancel 
                  Height          =   405
                  Left            =   3150
                  TabIndex        =   16
                  Top             =   180
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   714
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
                  ImgSize         =   32
                  cBack           =   -2147483633
                  mPointer        =   99
                  mIcon           =   "frm_CUSTOMER_DEPOSIT.frx":1725
               End
               Begin lvButton.lvButtons_H cmdChequeSave 
                  Height          =   405
                  Left            =   90
                  TabIndex        =   17
                  Top             =   180
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   714
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
                  ImgSize         =   32
                  Enabled         =   0   'False
                  cBack           =   -2147483633
                  mPointer        =   99
                  mIcon           =   "frm_CUSTOMER_DEPOSIT.frx":1A3F
               End
               Begin lvButton.lvButtons_H cmdCredit 
                  Height          =   405
                  Left            =   1620
                  TabIndex        =   18
                  Top             =   180
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   714
                  Caption         =   "&Credit"
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
                  ImgSize         =   32
                  Enabled         =   0   'False
                  cBack           =   -2147483633
                  mPointer        =   99
                  mIcon           =   "frm_CUSTOMER_DEPOSIT.frx":1D59
               End
            End
         End
         Begin OCX.b8Container b8Container1 
            Height          =   6525
            Left            =   0
            TabIndex        =   2
            Top             =   360
            Width           =   6285
            _ExtentX        =   11086
            _ExtentY        =   11509
            BackColor       =   16185592
            Begin VB.Frame fraAgencyDetails 
               BackColor       =   &H00F6F8F8&
               Height          =   2145
               Left            =   150
               TabIndex        =   6
               Top             =   90
               Width           =   6045
               Begin VB.TextBox txtCashCustomer 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1410
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   67
                  Text            =   " "
                  Top             =   300
                  Width           =   2985
               End
               Begin VB.TextBox txtCashPaidBy 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1410
                  MaxLength       =   100
                  TabIndex        =   9
                  Text            =   " "
                  Top             =   1500
                  Width           =   2985
               End
               Begin VB.TextBox txtCashAmount 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1410
                  MaxLength       =   20
                  TabIndex        =   8
                  Text            =   " "
                  Top             =   705
                  Width           =   2985
               End
               Begin MSComCtl2.DTPicker DTPCashDate 
                  Height          =   345
                  Left            =   1410
                  TabIndex        =   7
                  Top             =   1080
                  Width           =   3015
                  _ExtentX        =   5318
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
                  CustomFormat    =   "dd/MM/yyyy"
                  Format          =   58916867
                  CurrentDate     =   39257
               End
               Begin lvButton.lvButtons_H cmdCashCustomers 
                  Height          =   375
                  Left            =   4440
                  TabIndex        =   70
                  Top             =   270
                  Width           =   1515
                  _ExtentX        =   2672
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
                  Image           =   "frm_CUSTOMER_DEPOSIT.frx":2073
                  ImgSize         =   24
                  cBack           =   -2147483633
                  mPointer        =   99
                  mIcon           =   "frm_CUSTOMER_DEPOSIT.frx":2588
               End
               Begin VB.Label Label12 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Paid By"
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
                  Top             =   1500
                  Width           =   1485
               End
               Begin VB.Label Label8 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Amount"
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
                  TabIndex        =   12
                  Top             =   705
                  Width           =   1485
               End
               Begin VB.Label Label6 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Date Paid"
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
                  TabIndex        =   11
                  Top             =   1095
                  Width           =   1215
               End
               Begin VB.Label Label5 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Customers"
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
                  TabIndex        =   10
                  Top             =   300
                  Width           =   1215
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00F6F8F8&
               Height          =   675
               Left            =   150
               TabIndex        =   3
               Top             =   2220
               Width           =   6045
               Begin lvButton.lvButtons_H cmdCashCancel 
                  Height          =   405
                  Left            =   2940
                  TabIndex        =   4
                  Top             =   180
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   714
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
                  ImgSize         =   32
                  cBack           =   -2147483633
                  mPointer        =   99
                  mIcon           =   "frm_CUSTOMER_DEPOSIT.frx":28A2
               End
               Begin lvButton.lvButtons_H cmdCashSave 
                  Height          =   405
                  Left            =   1410
                  TabIndex        =   5
                  Top             =   180
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   714
                  Caption         =   "&Credit"
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
                  ImgSize         =   32
                  cBack           =   -2147483633
                  mPointer        =   99
                  mIcon           =   "frm_CUSTOMER_DEPOSIT.frx":2BBC
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frm_CUSTOMER_DEPOSIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtAmount_Change()

End Sub



Private Sub cboBankDepositBankName_Click()
    If cboBankDepositBankName.ListIndex = -1 Then Exit Sub
'    Call loadBankAccounts(cboBankDepositBankName.ItemData(cboBankDepositBankName.ListIndex), cboBankDepositAccountNo)
End Sub

Private Sub cboChequeBanks_Click()
    If cboChequeBanks.ListIndex = -1 Then Exit Sub
'    Call loadBankAccounts(cboChequeBanks.ItemData(cboChequeBanks.ListIndex), cboChequeAccount)
End Sub

Private Sub cboChequeStatus_Click()
    If cboChequeStatus.ListIndex = 0 Then
        cmdCredit.Enabled = True
        cmdChequeSave.Enabled = False
        Else
            cmdCredit.Enabled = False
            cmdChequeSave.Enabled = True
    End If
End Sub

Private Sub cmdBankDepositCancel_Click()
    Unload Me
End Sub

Private Sub cmdBankDepositCustomers_Click()
    blnCustomerDeposit = True
    frm_CUSTOMERS.Show
End Sub

Private Sub cmdBankDepositSave_Click()
'On Error GoTo errHandler
    
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtBankDepositCustomer, "Please select the customer name") Then Exit Sub
    If optBankDepositCash.Value = 0 Then
        If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtBankDepositChequeNo, "Please enter cheque number") Then Exit Sub
        If Mdl_FUNCTIONS.fn_CHECK_EMPTY_COMBO(cboBankDepositBankName, "Please select the bank name") Then Exit Sub
    End If
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtBankDepositAccountNo, "Please enter the account number") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtBankDepositAmount, "Please enter the amount") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtBankDeposited, "Please enter the bank name") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtBankDepositBranchCode, "Please enter the bank branch code") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtBankDepositTransactionNo, "Please enter the transaction number") Then Exit Sub
'    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_COMBO(cboBankDepositCurrency, "Please select the currency") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtBankDepositDepositedBy, "Please enter the name of the payee") Then Exit Sub
    
    
    With cls_BANK_DEPOSIT_Obj
    
        .CustomerID = lngSelectedCustomerID
        .ChequeNo = Trim(txtBankDepositChequeNo.Text)
        If cboBankDepositBankName.ListIndex = -1 Then
            .BankName = 0
            Else
            .BankName = cboBankDepositBankName.ItemData(cboBankDepositBankName.ListIndex)
        End If
        .Amount = Val(Trim(txtBankDepositAmount.Text))
        .BankDeposited = Trim(txtBankDeposited.Text)
        .BranchCode = Trim(txtBankDepositBranchCode.Text)
        .TransactionNo = Trim(txtBankDepositTransactionNo.Text)
        .AccountNo = Trim(txtBankDepositAccountNo.Text)
'        .CurrencyName = cboBankDepositCurrency.ItemData(cboBankDepositCurrency.ListIndex)
        .DatePaid = DTPBankDepositDatePaid.Value
        .DateReceived = DTPBankDepositDateReceived.Value
        .DepositedBy = Trim(txtBankDepositDepositedBy.Text)
        
        Call .fn_SAVE_BANK_DEPOSIT_RECORDS
        
        
    End With
    
    With cls_CASH_CHEQUE_BANK_DEPOSIT_Obj
        .TransactionID = 3
        .TransactionName = "Bank Deposit"
        .CustomerID = lngSelectedCustomerID
        .CustomerName = txtBankDepositCustomer.Text
        .ChequeNo = Trim(txtBankDepositChequeNo.Text)
        If cboBankDepositBankName.ListIndex = -1 Then
            .BankID = 0
            Else
                .BankID = cboBankDepositBankName.ItemData(cboBankDepositBankName.ListIndex)
        End If
        .BankName = cboBankDepositBankName.Text
        .AccountNo = Trim(txtBankDepositAccountNo.Text)
        .ChequeDate = Date
        .Amount = Val(Trim(txtBankDepositAmount.Text))
        .DateReceived = DTPBankDepositDateReceived.Value
        .PaidBy = Trim(txtBankDepositDepositedBy.Text)
        .ChequeStatus = ""
        .BankDeposited = Trim(txtBankDeposited.Text)
        .BranchCode = Trim(txtBankDepositBranchCode.Text)
        .TransactionNo = Trim(txtBankDepositTransactionNo.Text)
'        .CurrencyName = cboBankDepositCurrency.Text
        .fn_SAVE_DEPOSIT_RECORDS
        .fn_ADD_CUSTOMER_AMOUNT (lngSelectedCustomerID)
    End With
    
    MsgBox "Bank Deposit Details Saved Successfully", vbInformation, "Cheque Details"
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me)
    
'EXITPROCEDURE:
'    Exit Sub
'
'errHandler:
'    MsgBox Err.Description, vbCritical, "Connection"
'    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
'    GoTo EXITPROCEDURE
End Sub

Private Sub cmdCashCancel_Click()
    Unload Me
End Sub

Private Sub cmdCashCustomers_Click()
    blnCustomerDeposit = True
    frm_CUSTOMERS.Show
End Sub

Private Sub cmdCashSave_Click()
'On Error GoTo errHandler
    
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtCashCustomer, "Please enter the customer name") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtCashAmount, "Please enter amount") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtCashPaidBy, "Please enter the name of the payee") Then Exit Sub
    
    With cls_CASH_Obj
        .CustomerID = lngSelectedCustomerID
        .Amount = Trim(txtCashAmount.Text)
        .DatePaid = DTPCashDate.Value
        .PaidBy = Trim(txtCashPaidBy.Text)
        
        Call .fn_SAVE_CASH_RECORDS
    End With
    
    With cls_CASH_CHEQUE_BANK_DEPOSIT_Obj
        .TransactionID = 1
        .TransactionName = "Cash"
        .CustomerID = lngSelectedCustomerID
        .CustomerName = txtCashCustomer.Text
        .ChequeNo = ""
        .BankID = 0
        .BankName = ""
        .AccountNo = ""
        .ChequeDate = Date
        .Amount = Trim(txtCashAmount.Text)
        .DateReceived = DTPCashDate.Value
        .PaidBy = Trim(txtCashPaidBy.Text)
        .ChequeStatus = ""
        .BankDeposited = ""
        .BranchCode = ""
        .TransactionNo = ""
        .fn_SAVE_DEPOSIT_RECORDS
        .fn_ADD_CUSTOMER_AMOUNT (lngSelectedCustomerID)
    End With
    
    MsgBox "Cash Details Saved Successfully", vbInformation, "Cash Details"
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    
'EXITPROCEDURE:
'    Exit Sub
'
'errHandler:
'    MsgBox Err.Description, vbCritical, "Connection"
'    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
'    GoTo EXITPROCEDURE
End Sub

Private Sub cmdChequeCancel_Click()
    Unload Me
End Sub

Private Sub cmdChequeCustomers_Click()
    blnCustomerDeposit = True
    frm_CUSTOMERS.Show
End Sub

Private Sub cmdChequeSave_Click()
'On Error GoTo errHandler
    
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtChequeCustomer, "Please select the agency name") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtChequeChequeNo, "Please enter cheque number") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_COMBO(cboChequeBanks, "Please select the bank name") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtChequeAccountNo, "Please enter the account number") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtChequeAmount, "Please enter the amount") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtChequePaidBy, "Please enter the name of the payee") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_COMBO(cboChequeStatus, "Please select the status") Then Exit Sub
    
    Dim lngTransaction As Long
    lngTransaction = 2
    
    With cls_CHEQUE_Obj
    
        .CustomerID = cboChequeAgencyName.ItemData(cboChequeAgencyName.ListIndex)
        .ChequeNo = Trim(txtChequeChequeNo.Text)
        .BankName = cboChequeBanks.ItemData(cboChequeBanks.ListIndex)
        .AccountNo = Trim(txtChequeAccountNo.Text)
        .Amount = Val(Trim(txtChequeAmount.Text))
        .ChequeDate = DTPChequeChequeDate.Value
        .DateReceived = DTPChequeDateReceived.Value
        .PaidBy = Trim(txtChequePaidBy.Text)
        .Status = cboChequeStatus.ListIndex
        
        Call .fn_SAVE_CHECK_RECORDS
        
        
    End With
    
    With cls_CASH_CHEQUE_BANK_DEPOSIT_Obj
        .TransactionID = 2
        .TransactionName = "Cheque"
        .CustomerID = lngSelectedCustomerID
        .CustomerName = txtChequeCustomer.Text
        .ChequeNo = Trim(txtChequeChequeNo.Text)
        .BankID = cboChequeBanks.ItemData(cboChequeBanks.ListIndex)
        .BankName = cboChequeBanks.Text
        .AccountNo = Trim(txtChequeAccountNo.Text)
        .ChequeDate = DTPChequeChequeDate.Value
        .Amount = Val(Trim(txtChequeAmount.Text))
        .DateReceived = DTPChequeDateReceived.Value
        .PaidBy = Trim(txtChequePaidBy.Text)
        If cboChequeStatus.ListIndex = 0 Then
            .ChequeStatus = "Cleared"
            Else
                .ChequeStatus = "Not Cleared"
        End If
        .BankDeposited = ""
        .BranchCode = ""
        .TransactionNo = ""
        .CurrencyName = ""
        .fn_SAVE_DEPOSIT_RECORDS
        .fn_ADD_CUSTOMER_AMOUNT (lngSelectedCustomerID)
    End With
    
    MsgBox "Cheque Details Saved Successfully", vbInformation, "Cheque Details"
    
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me)
    
'EXITPROCEDURE:
'    Exit Sub
'
'errHandler:
'    MsgBox Err.Description, vbCritical, "Connection"
'    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
'    GoTo EXITPROCEDURE
End Sub

Private Sub cmdCredit_Click()
'On Error GoTo errHandler
    
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtChequeCustomer, "Please select the customer name") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtChequeChequeNo, "Please enter cheque number") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_COMBO(cboChequeBanks, "Please select the bank name") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtChequeAccountNo, "Please enter the account number") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtChequeAmount, "Please enter the amount") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtChequePaidBy, "Please enter the name of the payee") Then Exit Sub
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_COMBO(cboChequeStatus, "Please select the status") Then Exit Sub
    
    With cls_CHEQUE_Obj
    
        .CustomerID = lngSelectedCustomerID
        .ChequeNo = Trim(txtChequeChequeNo.Text)
        .BankName = cboChequeBanks.ItemData(cboChequeBanks.ListIndex)
        .AccountNo = Trim(txtChequeAccountNo.Text)
        .Amount = Val(Trim(txtChequeAmount.Text))
        .ChequeDate = DTPChequeChequeDate.Value
        .DateReceived = DTPChequeDateReceived.Value
        .PaidBy = Trim(txtChequePaidBy.Text)
        .Status = cboChequeStatus.ListIndex
        
        Call .fn_SAVE_CHECK_RECORDS
        
        
    End With
    
        With cls_CASH_CHEQUE_BANK_DEPOSIT_Obj
        .TransactionID = 2
        .TransactionName = "Cheque"
        .CustomerID = lngSelectedCustomerID
        .CustomerName = txtChequeCustomer.Text
        .ChequeNo = Trim(txtChequeChequeNo.Text)
        .BankID = cboChequeBanks.ItemData(cboChequeBanks.ListIndex)
        .BankName = cboChequeBanks.Text
        .AccountNo = Trim(txtChequeAccountNo.Text)
        .ChequeDate = DTPChequeChequeDate.Value
        .Amount = Val(Trim(txtChequeAmount.Text))
        .DateReceived = DTPChequeDateReceived.Value
        .PaidBy = Trim(txtChequePaidBy.Text)
        If cboChequeStatus.ListIndex = 0 Then
            .ChequeStatus = "Cleared"
            Else
                .ChequeStatus = "Not Cleared"
        End If
        .BankDeposited = ""
        .BranchCode = ""
        .TransactionNo = ""
        .CurrencyName = ""
        .fn_SAVE_DEPOSIT_RECORDS
        .fn_ADD_CUSTOMER_AMOUNT (lngSelectedCustomerID)
    End With
    
    MsgBox "Cheque Details Saved Successfully", vbInformation, "Cheque Details"
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me)
    
'EXITPROCEDURE:
'    Exit Sub
'
'errHandler:
'    MsgBox Err.Description, vbCritical, "Connection"
'    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
'    GoTo EXITPROCEDURE
End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_CENTER_FORM(Me)
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call subLoadBank(cboChequeBanks)
    Call subLoadBank(cboBankDepositBankName)
'    Call subLoadCurrency(cboBankDepositCurrency)
    
    SSTab1.Tab = 0
End Sub

Private Sub subLoadAgency(cbo As ComboBox)

    Dim rec As New ADODB.Recordset
    Set rec = cls_AGENCY_Obj.fn_LOAD_AGENCY(0)
    
    cbo.Clear
    
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            cbo.AddItem rec!AgencyName
            cbo.ItemData(cbo.NewIndex) = rec!AgencyID
            rec.MoveNext
        Loop
    End If
    
End Sub

Private Sub optBankDepositCash_Click()
    txtBankDepositChequeNo.Enabled = False
    cboBankDepositBankName.Enabled = False
End Sub

Private Sub optBankDepositCheque_Click()
    txtBankDepositChequeNo.Enabled = True
    cboBankDepositBankName.Enabled = True
End Sub



Private Sub txtBankDepositAmount_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtCashAmount_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtChequeAmount_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub subLoadCurrency(cbo As ComboBox)
'
'    Dim rec As New ADODB.Recordset
'    Set rec = cls_CURRENCY_Obj.fn_LOAD_CURRENCY(0)
'
'    cbo.Clear
'
'    If rec.AbsolutePosition <> -1 Then
'        Do While Not rec.EOF
'            cbo.AddItem rec!c_CurrencyName
'            cbo.ItemData(cbo.NewIndex) = rec!CurrencyID
'            rec.MoveNext
'        Loop
'    End If
    
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

Private Sub loadBankAccounts(lngID As Long, cbo As ComboBox)

    Dim rec As New ADODB.Recordset
    
    Set rec = cls_BANK_Obj.fn_LOAD_BANKS_ACCOUNTS(lngID)
    
    cbo.Clear
    
    If rec.AbsolutePosition = -1 Then Exit Sub
    Do While Not rec.EOF
        cbo.AddItem rec!AccountName
        cbo.ItemData(cbo.NewIndex) = rec!AccountID
        rec.MoveNext
    Loop

End Sub

Private Sub Save_CASH_CHEQUE_BANK_DEPOSIT(lngTransactionID As Long, lngAgencyID As Long, strAgencyName As String)
    
'    With cls_CASH_CHEQUE_BANK_DEPOSIT_Obj
'
'        .TransactionID = lngTransactionID
'        Select Case lngTransactionID
'            Case 1
'                .TransactionName = "Cash"
'            Case 2
'                .TransactionName = "Cheque"
'            Case 3
'                .TransactionName = "Bank Deposit"
'        End Select
'        .AgencyID = lngAgencyID
'        .AgencyName = strAgencyName
'        .ChequeNo
'
'        Call .fn_SAVE_DEPOSIT_RECORDS
'
'    End With
    

End Sub
