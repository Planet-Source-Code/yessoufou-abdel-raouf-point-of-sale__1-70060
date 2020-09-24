VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_SALARIES 
   BackColor       =   &H00FFFFFF&
   Caption         =   "EMPLOYEES SALARIES"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   12180
   Begin OCX.b8Container ContainerList 
      Height          =   7935
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   13996
      BackColor       =   16185592
      Begin OCX.b8Container b8Container4 
         Height          =   7725
         Left            =   2850
         TabIndex        =   14
         Top             =   120
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   13626
         BackColor       =   16185592
         Begin VB.Frame fraEmployeeDetails 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   7455
            Left            =   90
            TabIndex        =   15
            Top             =   90
            Width           =   6495
            Begin VB.Frame Frame7 
               Appearance      =   0  'Flat
               BackColor       =   &H00F6F8F8&
               Caption         =   "Salary Period"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   2655
               Left            =   60
               TabIndex        =   55
               Top             =   8040
               Width           =   6375
               Begin MSComctlLib.ListView lvw 
                  Height          =   2175
                  Left            =   150
                  TabIndex        =   56
                  Top             =   330
                  Width           =   6105
                  _ExtentX        =   10769
                  _ExtentY        =   3836
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
                  NumItems        =   3
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "ID"
                     Object.Width           =   0
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Text            =   "Start Date"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   2
                     Text            =   "End Date"
                     Object.Width           =   2540
                  EndProperty
               End
            End
            Begin VB.CheckBox ChkAlerte 
               BackColor       =   &H000000C0&
               Caption         =   "This employee is no more working here"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   345
               Left            =   30
               TabIndex        =   38
               Top             =   7080
               Visible         =   0   'False
               Width           =   6285
            End
            Begin VB.Frame fraGrossPay 
               Appearance      =   0  'Flat
               BackColor       =   &H00F6F8F8&
               Caption         =   "Gross Pay"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   765
               Left            =   60
               TabIndex        =   35
               Top             =   3150
               Width           =   6375
               Begin VB.Frame Frame5 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00F6F8F8&
                  BorderStyle     =   0  'None
                  Caption         =   "Frame5"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   495
                  Left            =   1260
                  TabIndex        =   36
                  Top             =   240
                  Width           =   4965
                  Begin VB.TextBox txtGrossPay 
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
                     IMEMode         =   3  'DISABLE
                     Left            =   0
                     MaxLength       =   100
                     TabIndex        =   11
                     Top             =   60
                     Width           =   4965
                  End
               End
               Begin VB.Label Label15 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Gross Pay"
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
                  Left            =   150
                  TabIndex        =   37
                  Top             =   300
                  Width           =   1185
               End
            End
            Begin VB.Frame fraNetPay 
               Appearance      =   0  'Flat
               BackColor       =   &H00F6F8F8&
               Caption         =   "Net Pay"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   765
               Left            =   60
               TabIndex        =   32
               Top             =   3930
               Width           =   6375
               Begin VB.Frame Frame6 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00F6F8F8&
                  BorderStyle     =   0  'None
                  Caption         =   "Frame5"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   435
                  Left            =   1200
                  TabIndex        =   33
                  Top             =   240
                  Width           =   5055
                  Begin VB.TextBox txtNetPay 
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
                     IMEMode         =   3  'DISABLE
                     Left            =   60
                     MaxLength       =   100
                     TabIndex        =   12
                     Top             =   30
                     Width           =   4965
                  End
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Net Pay"
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
                  Left            =   150
                  TabIndex        =   34
                  Top             =   270
                  Width           =   2265
               End
            End
            Begin VB.Frame Frame3 
               Appearance      =   0  'Flat
               BackColor       =   &H00F6F8F8&
               Caption         =   "Basic Salary"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   795
               Left            =   60
               TabIndex        =   30
               Top             =   0
               Width           =   6375
               Begin VB.TextBox txtBasicSalary 
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
                  IMEMode         =   3  'DISABLE
                  Left            =   1260
                  MaxLength       =   100
                  TabIndex        =   0
                  Top             =   300
                  Width           =   4965
               End
               Begin VB.Label Label12 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Basic Salary"
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
                  Left            =   120
                  TabIndex        =   31
                  Top             =   270
                  Width           =   1245
               End
            End
            Begin VB.Frame Frame2 
               Appearance      =   0  'Flat
               BackColor       =   &H00F6F8F8&
               Caption         =   "Deductions"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   2325
               Left            =   3300
               TabIndex        =   23
               Top             =   810
               Width           =   3135
               Begin VB.Frame Frame4 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00F6F8F8&
                  BorderStyle     =   0  'None
                  Caption         =   "Frame4"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   405
                  Left            =   120
                  TabIndex        =   24
                  Top             =   1860
                  Width           =   2955
                  Begin VB.TextBox txtTotalDeductions 
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
                     IMEMode         =   3  'DISABLE
                     Left            =   1080
                     MaxLength       =   100
                     TabIndex        =   10
                     Top             =   0
                     Width           =   1785
                  End
                  Begin VB.Label Label13 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Total"
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
                     Left            =   0
                     TabIndex        =   25
                     Top             =   0
                     Width           =   945
                  End
               End
               Begin VB.TextBox txtPTax 
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
                  IMEMode         =   3  'DISABLE
                  Left            =   1200
                  MaxLength       =   100
                  TabIndex        =   9
                  Top             =   1410
                  Width           =   1785
               End
               Begin VB.TextBox txtIncomeTax 
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
                  IMEMode         =   3  'DISABLE
                  Left            =   1200
                  MaxLength       =   100
                  TabIndex        =   8
                  Top             =   1020
                  Width           =   1785
               End
               Begin VB.TextBox txtInssurance 
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
                  IMEMode         =   3  'DISABLE
                  Left            =   1200
                  MaxLength       =   100
                  TabIndex        =   7
                  Top             =   630
                  Width           =   1785
               End
               Begin VB.TextBox txtGPF 
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
                  IMEMode         =   3  'DISABLE
                  Left            =   1200
                  MaxLength       =   100
                  TabIndex        =   6
                  Top             =   240
                  Width           =   1785
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "P.Tax"
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
                  Left            =   120
                  TabIndex        =   29
                  Top             =   1410
                  Width           =   1095
               End
               Begin VB.Label Label9 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Income Tax"
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
                  Left            =   120
                  TabIndex        =   28
                  Top             =   1020
                  Width           =   1095
               End
               Begin VB.Label Label8 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Inssurance"
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
                  Left            =   120
                  TabIndex        =   27
                  Top             =   630
                  Width           =   1095
               End
               Begin VB.Label Label7 
                  BackStyle       =   0  'Transparent
                  Caption         =   "G.P.F"
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
                  Left            =   120
                  TabIndex        =   26
                  Top             =   240
                  Width           =   1095
               End
            End
            Begin VB.Frame Frame1 
               Appearance      =   0  'Flat
               BackColor       =   &H00F6F8F8&
               Caption         =   "Allowance"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   2325
               Left            =   60
               TabIndex        =   16
               Top             =   810
               Width           =   3135
               Begin VB.Frame fraTotal 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00F6F8F8&
                  BorderStyle     =   0  'None
                  Caption         =   "Frame4"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   525
                  Left            =   90
                  TabIndex        =   17
                  Top             =   1740
                  Width           =   2955
                  Begin VB.TextBox txtTotalAllowance 
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
                     IMEMode         =   3  'DISABLE
                     Left            =   1170
                     MaxLength       =   100
                     TabIndex        =   5
                     Top             =   120
                     Width           =   1785
                  End
                  Begin VB.Label Label11 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Total"
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
                     Left            =   30
                     TabIndex        =   18
                     Top             =   120
                     Width           =   945
                  End
               End
               Begin VB.TextBox txtTransport 
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
                  IMEMode         =   3  'DISABLE
                  Left            =   1260
                  MaxLength       =   100
                  TabIndex        =   4
                  Top             =   1410
                  Width           =   1785
               End
               Begin VB.TextBox txtCCA 
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
                  IMEMode         =   3  'DISABLE
                  Left            =   1260
                  MaxLength       =   100
                  TabIndex        =   3
                  Top             =   1020
                  Width           =   1785
               End
               Begin VB.TextBox txtHRA 
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
                  IMEMode         =   3  'DISABLE
                  Left            =   1260
                  MaxLength       =   100
                  TabIndex        =   2
                  Top             =   630
                  Width           =   1785
               End
               Begin VB.TextBox txtDA 
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
                  IMEMode         =   3  'DISABLE
                  Left            =   1260
                  MaxLength       =   100
                  TabIndex        =   1
                  Top             =   240
                  Width           =   1785
               End
               Begin VB.Label Label6 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Transport"
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
                  Left            =   120
                  TabIndex        =   22
                  Top             =   1410
                  Width           =   1095
               End
               Begin VB.Label Label5 
                  BackStyle       =   0  'Transparent
                  Caption         =   "C.C.A"
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
                  Left            =   120
                  TabIndex        =   21
                  Top             =   1020
                  Width           =   1095
               End
               Begin VB.Label Label4 
                  BackStyle       =   0  'Transparent
                  Caption         =   "H.R.A"
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
                  Left            =   120
                  TabIndex        =   20
                  Top             =   630
                  Width           =   1095
               End
               Begin VB.Label Label3 
                  BackStyle       =   0  'Transparent
                  Caption         =   "D.A"
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
                  Left            =   120
                  TabIndex        =   19
                  Top             =   240
                  Width           =   1095
               End
            End
         End
      End
      Begin OCX.b8Container b8Container3 
         Height          =   7725
         Left            =   60
         TabIndex        =   39
         Top             =   120
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   13626
         BackColor       =   16185592
         Begin VB.ListBox lstEmployees 
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
            Height          =   7050
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   40
            Top             =   480
            Width           =   2565
         End
         Begin OCX.b8SideTab b8SideTab2 
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   150
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   661
            Caption         =   "Employees"
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
      Begin OCX.b8Container PicContainer 
         Height          =   7725
         Left            =   9570
         TabIndex        =   42
         Top             =   120
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   13626
         BackColor       =   16185592
         Begin VB.Frame fraPicture 
            Appearance      =   0  'Flat
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   7215
            Left            =   90
            TabIndex        =   43
            Top             =   60
            Width           =   2145
            Begin VB.PictureBox Picture1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   2685
               Left            =   30
               ScaleHeight     =   2655
               ScaleWidth      =   2055
               TabIndex        =   44
               Top             =   90
               Width           =   2085
               Begin VB.Image imgPicture 
                  Height          =   2655
                  Left            =   0
                  Picture         =   "frm_SALARIES.frx":0000
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   2055
               End
               Begin VB.Label lblAlerte 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "No Picture"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Left            =   60
                  TabIndex        =   45
                  Top             =   1110
                  Width           =   1965
               End
            End
            Begin lvButton.lvButtons_H cmdIdentification 
               Height          =   375
               Left            =   30
               TabIndex        =   46
               Top             =   2790
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   661
               CapAlign        =   2
               BackStyle       =   3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFore           =   16777215
               cFHover         =   16777215
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16711680
            End
         End
      End
   End
   Begin OCX.b8Container b8Container2 
      Height          =   675
      Left            =   0
      TabIndex        =   47
      Top             =   7950
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   1191
      BackColor       =   16185592
      Begin lvButton.lvButtons_H cmdAddNew 
         Height          =   495
         Left            =   180
         TabIndex        =   48
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "&Add New"
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
         Image           =   "frm_SALARIES.frx":26B8
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SALARIES.frx":2BCD
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   7000
         TabIndex        =   49
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
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
         Image           =   "frm_SALARIES.frx":2EE7
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SALARIES.frx":3456
      End
      Begin lvButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   1885
         TabIndex        =   50
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
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
         Image           =   "frm_SALARIES.frx":3770
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SALARIES.frx":3AED
      End
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   495
         Left            =   10410
         TabIndex        =   51
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
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
         Image           =   "frm_SALARIES.frx":3E07
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SALARIES.frx":4397
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   5295
         TabIndex        =   52
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
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
         Image           =   "frm_SALARIES.frx":46B1
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SALARIES.frx":4911
      End
      Begin lvButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   3590
         TabIndex        =   53
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
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
         Image           =   "frm_SALARIES.frx":4C2B
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SALARIES.frx":513E
      End
      Begin lvButton.lvButtons_H cmdPrint 
         Height          =   495
         Left            =   8705
         TabIndex        =   54
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "&Pay Salary"
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
         Image           =   "frm_SALARIES.frx":5458
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SALARIES.frx":5772
      End
   End
End
Attribute VB_Name = "frm_SALARIES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngID As Long
Public lngSalaryID As Long

Public lngSalaryPeriodID As Long

Dim blnSalaryModify As Boolean
Dim blnSalaryAdd As Boolean

Private Sub cmdCancel_Click()
    Call fn_DISABLE_CONTROLS
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    fraEmployeeDetails.Enabled = False
    lstEmployees.Enabled = True
    
    fraGrossPay.Visible = True
    fraNetPay.Visible = True
    
    cmdDelete.Enabled = False
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub



Private Sub cmdaddNew_Click()
        
    lngSalaryID = cls_SALARIES_Obj.fn_AUTOGEN
        
    blnSalaryAdd = True
    blnSalaryModify = False
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call fn_ENABLE_CONTROLS
    Call fn_UNSET_CONTROL_COLOR(Me, "All")
    fraEmployeeDetails.Enabled = True
    lstEmployees.Enabled = False
    
'    fraGrossPay.Visible = False
'    fraNetPay.Visible = False
    txtBasicSalary.SetFocus
    
End Sub

Private Sub cmdPrint_Click()
    
    frm_SALARY_PERIOD.Show
    
    If lngSalaryPeriodID <= 0 Then
        Exit Sub
    End If
    
    With cls_SALARIES_Obj
        .SalaryID = lngSalaryID
        .SalaryPeriodID = lngSalaryPeriodID
        .fn_SAVE_SALARY_PERIOD_DETAILS
    End With

    With frm_PAY_SLIP
        .sub_LOAD_REPORT (lngID)
        .Show
    End With
    
End Sub

Private Sub cmdSave_Click()

    txtGrossPay.Text = Val(txtBasicSalary.Text) + Val(txtTotalAllowance.Text)
    txtNetPay.Text = Val(txtGrossPay.Text) - Val(txtTotalDeductions.Text)
    
    
'    If MdlFunctions.fnRequireTextField(txtCategoryName, "Kindly enter category name.") Then GoTo EXITPROCEDURE
    
    With cls_SALARIES_Obj
    
        .BasicPay = Val(txtBasicSalary.Text)
        .TotalAllowance = Val(txtTotalAllowance.Text)
        .GrossPay = Val(txtGrossPay.Text)
        .TotalDeduction = Val(txtTotalDeductions.Text)
        .NetPay = Val(txtNetPay.Text)
        
        If lngSalaryID = 0 Then
            .SalaryID = .fn_AUTOGEN
            Else
                .SalaryID = lngSalaryID
        End If
        
        .DA = Val(txtDA.Text)
        .HRA = Val(txtHRA.Text)
        .CCA = Val(txtCCA.Text)
        .Transport = Val(txtTransport.Text)
        
        .GPF = Val(txtGPF.Text)
        .Inssurance = Val(txtGPF.Text)
        .IncomeTax = Val(txtIncomeTax.Text)
        .PTax = Val(txtPTax.Text)
        
    

        If blnSalaryAdd = True And blnSalaryModify = False Then
            Call .fn_SAVE_SALARIES(lngSalaryID, lngID)
            Call .fn_SAVE_ALLOWANCE_DETAILS(lngSalaryID)
            Call .fn_SAVE_DEDUCTION_DETAILS(lngSalaryID)
            MsgBox "Salary successfully saved.", vbExclamation, Title
            Else
                .fn_UPDATE_SALARIES (lngID)
                .fn_UPDATE_ALLOWANCE_DETAILS (lngSalaryID)
                .fn_UPDATE_DEDUCTION_DETAILS (lngSalaryID)
                MsgBox "Salary successfully updated.", vbExclamation, Title
        End If
        
        
    End With

    
    fraGrossPay.Visible = True
    fraNetPay.Visible = True
    
    Call fn_DISABLE_CONTROLS
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    fraEmployeeDetails.Enabled = False
    fraPicture.Enabled = False
'    Call subLoadSalaryDetails(lstEmployees)
    lstEmployees.Enabled = True
    
    cmdDelete.Enabled = False
    
EXITPROCEDURE:
    Exit Sub

End Sub

Private Sub cmdEdit_Click()

    blnSalaryAdd = False
    blnSalaryModify = True
    Call fn_ENABLE_CONTROLS
    Call Mdl_FUNCTIONS.fn_UNSET_CONTROL_COLOR(Me, "All")
    fraEmployeeDetails.Enabled = True
    lstEmployees.Enabled = False

'    fraGrossPay.Visible = False
'    fraNetPay.Visible = False
    txtBasicSalary.SetFocus
End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_FORM_SIZE(Me)
    Move 0, 0
    Call fn_SET_CONTROL_COLOR(Me, "All")
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call fn_DISABLE_CONTROLS
    Call sub_LOAD_EMPLOYEES(lstEmployees)

    
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdAddNew, cmdAddNew.hwnd, "Add new salary details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdEdit, cmdEdit.hwnd, "Edit existing salary details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdDelete, cmdDelete.hwnd, "Delete existing salary details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSave, cmdSave.hwnd, "Save salary details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancel, cmdCancel.hwnd, "Cancel the transaction.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdClose, cmdClose.hwnd, "Close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor

    cmdDelete.Enabled = False

End Sub

Private Sub sub_LOAD_EMPLOYEES(Optional lst As ListBox)

    Dim rec As New ADODB.Recordset
    Set rec = cls_EMPLOYEES_Obj.fn_LOAD_EMPLOYEES
    
    lst.Clear
    
    Do While Not rec.EOF
        lst.AddItem rec!FirstName & " " & rec!LastName
        lst.ItemData(lst.NewIndex) = rec!EmployeeID
        rec.MoveNext
    Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)
    blnSalaryAdd = False
    blnSalaryModify = False
End Sub

Private Sub lstEmployees_Click()
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    lngID = lstEmployees.ItemData(lstEmployees.ListIndex)
    Call sub_LOAD_EMPLOYEE_PICTURE(lngID)
    Call sub_LOAD_SALARY_DETAILS(lngID)
    
    If txtBasicSalary.Text <> "" Then
        cmdAddNew.Enabled = False
        cmdEdit.Enabled = True
        Else
            cmdAddNew.Enabled = True
            cmdEdit.Enabled = False
    End If
    
End Sub

Private Sub sub_LOAD_SALARY_DETAILS(lngEmployeeID As Long)

    Dim rec As New ADODB.Recordset
    
    Set rec = cls_SALARIES_Obj.fn_LOAD_SALARIES(lngEmployeeID)
      
    If rec.AbsolutePosition <> -1 Then
    
        lngSalaryID = rec!SalaryID
        txtBasicSalary.Text = rec!BasicPay
        txtTotalAllowance.Text = rec!TotalAllowance
        txtDA.Text = rec!DA
        txtHRA.Text = rec!HRA
        txtCCA.Text = rec!CCA
        txtTransport.Text = rec!Transport
        txtGrossPay.Text = rec!GrossPay
        txtGPF.Text = rec!GPF
        txtInssurance.Text = rec!Inssurance
        txtIncomeTax.Text = rec!IncomeTax
        txtPTax.Text = rec!PTax
        txtTotalDeductions.Text = rec!TotalDeduction
        txtNetPay.Text = rec!NetPay

    End If

    
End Sub

Private Sub sub_LOAD_EMPLOYEE_PICTURE(lngEmployeeID As Long)
On Error GoTo errHandler
    Dim rec As New ADODB.Recordset
    
    Set rec = cls_EMPLOYEES_Obj.fn_LOAD_EMPLOYEES(lngEmployeeID)
    
    cmdIdentification.Caption = rec!EmployeeNo
    
        If rec!WorkingStatus = 1 Then
        ChkAlerte.Visible = True
        ChkAlerte.Value = 1
        Else
            ChkAlerte.Visible = False
    End If
    
    If rec!Photo = "" Or IsNull(rec!Photo) Then
        imgPicture.Picture = LoadPicture()
        Else
            imgPicture.Picture = LoadPicture(strPicturePath & rec!Photo)
    End If
    
EXITPROCEDURE:
    Exit Sub
    
errHandler:
    If Err.Number = 53 Then
        imgPicture.Picture = LoadPicture()
        Exit Sub
    End If
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, Me.Name, "cmdTelecharger")
    GoTo EXITPROCEDURE
End Sub

Private Sub txtBasicSalary_Change()
'    txtGrossPay.Text = Val(txtBasicSalary.Text) + Val(txtTotalAllowance.Text)
'    txtNetPay.Text = Val(txtBasicSalary.Text) + Val(txtTotalAllowance.Text) - Val(txtTotalDeductions.Text)
End Sub

Private Sub txtBasicSalary_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtCCA_Change()
    txtTotalAllowance.Text = Val(txtDA.Text) + Val(txtHRA.Text) + Val(txtCCA.Text) + Val(txtTransport.Text)
'    txtGrossPay.Text = Val(txtBasicSalary.Text) + Val(txtTotalAllowance.Text)

End Sub

Private Sub txtCCA_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtDA_Change()
    txtTotalAllowance.Text = Val(txtDA.Text) + Val(txtHRA.Text) + Val(txtCCA.Text) + Val(txtTransport.Text)
'    txtGrossPay.Text = Val(txtBasicSalary.Text) + Val(txtTotalAllowance.Text)
End Sub

Private Sub txtDA_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtGPF_Change()
    txtTotalDeductions.Text = Val(txtGPF.Text) + Val(txtInssurance.Text) + Val(txtIncomeTax.Text) + Val(txtPTax.Text)
'    txtNetPay.Text = Val(txtGrossPay.Text) - Val(txtTotalDeductions.Text)
End Sub

Private Sub txtGPF_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtHRA_Change()
    txtTotalAllowance.Text = Val(txtDA.Text) + Val(txtHRA.Text) + Val(txtCCA.Text) + Val(txtTransport.Text)
'    txtGrossPay.Text = Val(txtBasicSalary.Text) + Val(txtTotalAllowance.Text)

End Sub

Private Sub txtHRA_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtIncomeTax_Change()
    txtTotalDeductions.Text = Val(txtGPF.Text) + Val(txtInssurance.Text) + Val(txtIncomeTax.Text) + Val(txtPTax.Text)
'    txtNetPay.Text = Val(txtGrossPay.Text) - Val(txtTotalDeductions.Text)
End Sub

Private Sub txtIncomeTax_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtInssurance_Change()
    txtTotalDeductions.Text = Val(txtGPF.Text) + Val(txtInssurance.Text) + Val(txtIncomeTax.Text) + Val(txtPTax.Text)
'    txtNetPay.Text = Val(txtGrossPay.Text) - Val(txtTotalDeductions.Text)
End Sub

Private Sub txtInssurance_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtPTax_Change()
    txtTotalDeductions.Text = Val(txtGPF.Text) + Val(txtInssurance.Text) + Val(txtIncomeTax.Text) + Val(txtPTax.Text)
'    txtNetPay.Text = Val(txtGrossPay.Text) - Val(txtTotalDeductions.Text)
End Sub

Private Sub txtPTax_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtTransport_Change()
    txtTotalAllowance.Text = Val(txtDA.Text) + Val(txtHRA.Text) + Val(txtCCA.Text) + Val(txtTransport.Text)
'    txtGrossPay.Text = Val(txtBasicSalary.Text) + Val(txtTotalAllowance.Text)

End Sub

Private Sub txtTransport_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Public Function fn_DISABLE_CONTROLS()

    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdAddNew.Enabled = True

'    cmdLoadPicture.Enabled = False
'    cmdRemovePicture.Enabled = False

End Function

'**********Function to enable some buttons***********
Public Function fn_ENABLE_CONTROLS()

    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdAddNew.Enabled = False
'
'    cmdLoadPicture.Enabled = True
'    cmdRemovePicture.Enabled = True


End Function
