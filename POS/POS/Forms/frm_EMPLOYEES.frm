VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_EMPLOYEES 
   BackColor       =   &H00FFFFFF&
   Caption         =   "EMPLOYEES"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   12180
   Begin OCX.b8Container ContainerList 
      Height          =   7755
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   13679
      BackColor       =   16185592
      Begin OCX.b8Container b8Container4 
         Height          =   7575
         Left            =   3810
         TabIndex        =   16
         Top             =   90
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   13361
         BackColor       =   16185592
         Begin VB.Frame fraEmployeeDetails 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   7245
            Left            =   90
            TabIndex        =   17
            Top             =   60
            Width           =   5385
            Begin VB.Frame Frame1 
               Appearance      =   0  'Flat
               BackColor       =   &H00F6F8F8&
               Caption         =   "Company Info"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   2805
               Left            =   60
               TabIndex        =   18
               Top             =   -30
               Width           =   5205
               Begin VB.ComboBox cboReportsTo 
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
                  Left            =   1680
                  Style           =   2  'Dropdown List
                  TabIndex        =   3
                  Top             =   1470
                  Width           =   3405
               End
               Begin VB.TextBox txtLastName 
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
                  Left            =   1680
                  MaxLength       =   100
                  TabIndex        =   1
                  Top             =   660
                  Width           =   3405
               End
               Begin VB.TextBox txtFirstName 
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
                  Left            =   1680
                  MaxLength       =   100
                  TabIndex        =   0
                  Top             =   270
                  Width           =   3405
               End
               Begin VB.TextBox txtTitle 
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
                  Left            =   1680
                  MaxLength       =   100
                  TabIndex        =   2
                  Top             =   1080
                  Width           =   3405
               End
               Begin VB.TextBox txtExtension 
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
                  Left            =   1680
                  MaxLength       =   100
                  TabIndex        =   5
                  Top             =   2370
                  Width           =   3375
               End
               Begin MSComCtl2.DTPicker DTHireDate 
                  Height          =   405
                  Left            =   1680
                  TabIndex        =   4
                  Top             =   1890
                  Width           =   3405
                  _ExtentX        =   6006
                  _ExtentY        =   714
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
                  Format          =   20643841
                  CurrentDate     =   39095
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Reports To"
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
                  Left            =   90
                  TabIndex        =   24
                  Top             =   1470
                  Width           =   960
               End
               Begin VB.Label Label2 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Last Name"
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
                  TabIndex        =   23
                  Top             =   660
                  Width           =   2265
               End
               Begin VB.Label Label12 
                  BackStyle       =   0  'Transparent
                  Caption         =   "First Name"
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
                  TabIndex        =   22
                  Top             =   270
                  Width           =   2265
               End
               Begin VB.Label Label13 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Title(Position)"
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
                  TabIndex        =   21
                  Top             =   1080
                  Width           =   2265
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hire Date"
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
                  Left            =   90
                  TabIndex        =   20
                  Top             =   1860
                  Width           =   825
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Extension"
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
                  Left            =   90
                  TabIndex        =   19
                  Top             =   2340
                  Width           =   840
               End
            End
            Begin VB.Frame Frame2 
               Appearance      =   0  'Flat
               BackColor       =   &H00F6F8F8&
               Caption         =   "Personnal Info"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   4455
               Left            =   60
               TabIndex        =   25
               Top             =   2760
               Width           =   5205
               Begin VB.TextBox txtNotes 
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
                  Height          =   645
                  IMEMode         =   3  'DISABLE
                  Left            =   1680
                  MaxLength       =   100
                  MultiLine       =   -1  'True
                  TabIndex        =   14
                  Top             =   3270
                  Width           =   3375
               End
               Begin VB.TextBox txtHomePhone 
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
                  Left            =   1680
                  MaxLength       =   50
                  TabIndex        =   13
                  Top             =   2880
                  Width           =   3375
               End
               Begin VB.TextBox txtCountry 
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
                  Left            =   1680
                  MaxLength       =   50
                  TabIndex        =   12
                  Top             =   2520
                  Width           =   3375
               End
               Begin VB.TextBox txtPostalCode 
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
                  Left            =   1680
                  MaxLength       =   50
                  TabIndex        =   11
                  Top             =   2160
                  Width           =   3375
               End
               Begin VB.TextBox txtRegion 
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
                  Left            =   1680
                  MaxLength       =   100
                  TabIndex        =   10
                  Top             =   1800
                  Width           =   3375
               End
               Begin VB.TextBox txtCity 
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
                  Left            =   1680
                  MaxLength       =   100
                  TabIndex        =   9
                  Top             =   1440
                  Width           =   3375
               End
               Begin VB.TextBox txtAddress 
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
                  Left            =   1680
                  MaxLength       =   200
                  TabIndex        =   8
                  Top             =   1080
                  Width           =   3375
               End
               Begin VB.ComboBox cboTitleOfCourtesy 
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
                  Left            =   1680
                  Style           =   2  'Dropdown List
                  TabIndex        =   6
                  Top             =   270
                  Width           =   3405
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
                  Left            =   120
                  TabIndex        =   26
                  Top             =   4020
                  Visible         =   0   'False
                  Width           =   4935
               End
               Begin MSComCtl2.DTPicker DTBirthDate 
                  Height          =   405
                  Left            =   1680
                  TabIndex        =   7
                  Top             =   630
                  Width           =   3405
                  _ExtentX        =   6006
                  _ExtentY        =   714
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
                  Format          =   20643841
                  CurrentDate     =   39095
               End
               Begin VB.Label Label11 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Notes"
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
                  TabIndex        =   35
                  Top             =   3180
                  Width           =   1125
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Home Phone"
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
                  TabIndex        =   34
                  Top             =   2880
                  Width           =   1215
               End
               Begin VB.Label Label9 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Country"
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
                  TabIndex        =   33
                  Top             =   2520
                  Width           =   1125
               End
               Begin VB.Label Label8 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Postal Code"
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
                  TabIndex        =   32
                  Top             =   2160
                  Width           =   1275
               End
               Begin VB.Label Label7 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Region"
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
                  TabIndex        =   31
                  Top             =   1800
                  Width           =   1035
               End
               Begin VB.Label Label6 
                  BackStyle       =   0  'Transparent
                  Caption         =   "City"
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
                  TabIndex        =   30
                  Top             =   1440
                  Width           =   1095
               End
               Begin VB.Label Label5 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address"
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
                  TabIndex        =   29
                  Top             =   1080
                  Width           =   1065
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Title Of Courtesy"
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
                  Left            =   90
                  TabIndex        =   28
                  Top             =   270
                  Width           =   1440
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Birth Date"
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
                  Left            =   90
                  TabIndex        =   27
                  Top             =   600
                  Width           =   870
               End
            End
         End
      End
      Begin OCX.b8Container b8Container3 
         Height          =   7575
         Left            =   60
         TabIndex        =   36
         Top             =   90
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   13361
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
            Height          =   6075
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   37
            Top             =   480
            Width           =   3465
         End
         Begin OCX.b8SideTab b8SideTab2 
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   150
            Width           =   3465
            _ExtentX        =   6112
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
         Begin lvButton.lvButtons_H cmdPrevious 
            Height          =   375
            Left            =   1020
            TabIndex        =   39
            Top             =   6690
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   661
            Caption         =   "7"
            CapAlign        =   2
            BackStyle       =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Webdings"
               Size            =   12
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EMPLOYEES.frx":0000
         End
         Begin lvButton.lvButtons_H cmdFirst 
            Height          =   375
            Left            =   150
            TabIndex        =   40
            Top             =   6690
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   661
            Caption         =   "9"
            CapAlign        =   2
            BackStyle       =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Webdings"
               Size            =   12
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EMPLOYEES.frx":031A
         End
         Begin lvButton.lvButtons_H cmdLast 
            Height          =   375
            Left            =   2760
            TabIndex        =   41
            Top             =   6690
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   661
            Caption         =   ":"
            CapAlign        =   2
            BackStyle       =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Webdings"
               Size            =   12
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EMPLOYEES.frx":0634
         End
         Begin lvButton.lvButtons_H cmdNext 
            Height          =   375
            Left            =   1890
            TabIndex        =   42
            Top             =   6690
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   661
            Caption         =   "8"
            CapAlign        =   2
            BackStyle       =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Webdings"
               Size            =   12
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EMPLOYEES.frx":094E
         End
      End
      Begin OCX.b8Container PicContainer 
         Height          =   7575
         Left            =   9420
         TabIndex        =   43
         Top             =   90
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   13361
         BackColor       =   16185592
         Begin VB.Frame fraPicture 
            Appearance      =   0  'Flat
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   7215
            Left            =   90
            TabIndex        =   44
            Top             =   60
            Width           =   2385
            Begin VB.PictureBox Picture1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   2685
               Left            =   60
               ScaleHeight     =   2655
               ScaleWidth      =   2205
               TabIndex        =   45
               Top             =   90
               Width           =   2235
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
                  TabIndex        =   46
                  Top             =   1110
                  Width           =   1965
               End
               Begin VB.Image imgPicture 
                  Height          =   2655
                  Left            =   0
                  Picture         =   "frm_EMPLOYEES.frx":0C68
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   2205
               End
            End
            Begin lvButton.lvButtons_H cmdLoadPicture 
               Height          =   375
               Left            =   60
               TabIndex        =   47
               Top             =   3300
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   661
               Caption         =   "Add/Change"
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
               cBack           =   -2147483633
               mPointer        =   99
               mIcon           =   "frm_EMPLOYEES.frx":3320
            End
            Begin lvButton.lvButtons_H cmdIdentification 
               Height          =   375
               Left            =   60
               TabIndex        =   48
               Top             =   2790
               Width           =   2250
               _ExtentX        =   3969
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
            Begin lvButton.lvButtons_H cmdRemovePicture 
               Height          =   375
               Left            =   60
               TabIndex        =   49
               Top             =   3810
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   661
               Caption         =   "Remove"
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
               cBack           =   -2147483633
               mPointer        =   99
               mIcon           =   "frm_EMPLOYEES.frx":363A
            End
         End
         Begin MSComDlg.CommonDialog PictureDlg 
            Left            =   210
            Top             =   4380
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
   End
   Begin OCX.b8Container b8Container2 
      Height          =   675
      Left            =   0
      TabIndex        =   50
      Top             =   7770
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   1191
      BackColor       =   16185592
      Begin lvButton.lvButtons_H cmdAddNew 
         Height          =   495
         Left            =   180
         TabIndex        =   51
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
         Image           =   "frm_EMPLOYEES.frx":3954
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_EMPLOYEES.frx":3E69
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   8196
         TabIndex        =   52
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
         Image           =   "frm_EMPLOYEES.frx":4183
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_EMPLOYEES.frx":46F2
      End
      Begin lvButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   2184
         TabIndex        =   53
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
         Image           =   "frm_EMPLOYEES.frx":4A0C
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_EMPLOYEES.frx":4D89
      End
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   495
         Left            =   10200
         TabIndex        =   54
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
         Image           =   "frm_EMPLOYEES.frx":50A3
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_EMPLOYEES.frx":5633
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   6192
         TabIndex        =   55
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
         Image           =   "frm_EMPLOYEES.frx":594D
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_EMPLOYEES.frx":5BAD
      End
      Begin lvButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   4188
         TabIndex        =   56
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
         Image           =   "frm_EMPLOYEES.frx":5EC7
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_EMPLOYEES.frx":63DA
      End
   End
End
Attribute VB_Name = "frm_EMPLOYEES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lngID As Long

Dim strPictureName As String

Dim blnEmployeeModify As Boolean
Dim blnEmployeeAdd As Boolean



Private Sub ChkAlerte_Click()
    If ChkAlerte.Value = 1 Then
    
        ChkAlerte.Caption = "This employee is no more working here"
        ChkAlerte.BackColor = &HC0&
        
        Else
        
            ChkAlerte.Caption = "This employee should continue working?"
            ChkAlerte.BackColor = &HC000&
        
    End If
End Sub

Private Sub cmdCancel_Click()
    Call fn_DISABLE_CONTROLS
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    fraEmployeeDetails.Enabled = False
    fraPicture.Enabled = False
    lstEmployees.Enabled = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
'On Error GoTo errHandler

    If txtFirstName.Text = "" Then
        MsgBox "Kindly choose the Employee to be Deleted.", vbExclamation, Title
        GoTo EXITPROCEDURE
    End If
    
    If MsgBox("Are you sure you want to delete  " & Trim(txtFirstName.Text) & " ?", vbQuestion + vbYesNo, Title) = vbNo Then
        
        GoTo EXITPROCEDURE
        
        Else
        
            Call cls_EMPLOYEES_Obj.fn_DELETE_EMPLOYEE_RECORDS(lngID)
            MsgBox "Employee successfully deleted.", vbExclamation, Title
            
    End If
    
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call sub_LOAD_TITLES(cboTitleOfCourtesy)
    Call sub_LOAD_SUPERVISORS(cboReportsTo)
    Call sub_LOAD_EMPLOYEES(lstEmployees)
'    Call subLoadCustomerDetails(lngID)
    
EXITPROCEDURE:
    Exit Sub
    
'errHandler:
'    MsgBox "An error occured.", vbCritical, title
'    Call MdlFunctions.fnWriteErrorToFile(Date, Time, Err.Description, Err.Number, Me.Name, "cmdCategorySupprimer")
'    GoTo EXITPROCEDURE
End Sub

Private Sub cmdaddNew_Click()
    blnEmployeeAdd = True
    blnEmployeeModify = False
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call fn_ENABLE_CONTROLS
    cmdIdentification.Caption = cls_EMPLOYEES_Obj.fn_AUTOGEN
    Call fn_UNSET_CONTROL_COLOR(Me, "All")
    fraEmployeeDetails.Enabled = True
    fraPicture.Enabled = True
    lstEmployees.Enabled = False
    txtFirstName.SetFocus
End Sub

Private Sub cmdLoadPicture_Click()
On Error GoTo errHandler
    
'    imgPicture.Picture = LoadPicture()
    PictureDlg.ShowOpen
    If PictureDlg.FileName = "" Then GoTo EXITPROCEDURE
    imgPicture.Picture = LoadPicture(PictureDlg.FileName)

EXITPROCEDURE:
    Exit Sub
    
errHandler:
    MsgBox "La photo n'est pas valide", vbCritical, "Connection"
    Call MdlFunctions.fnWriteErrorToFile(Date, Time, Err.Description, Err.Number, Me.Name, "cmdTelecharger")
    GoTo EXITPROCEDURE
End Sub

Private Sub cmdRemovePicture_Click()
    If MsgBox("Are you sure you want to remove  " & Trim(txtFirstName.Text) & "'s picture ?", vbQuestion + vbYesNo, Title) = vbNo Then
        
        GoTo EXITPROCEDURE
        
        Else
        
            Call cls_EMPLOYEES_Obj.fn_DELETE_EMPLOYEE_PICTURE(lngID)
            strPictureName = Trim(cmdIdentification.Caption) & ".bmp"
            Kill strPicturePath & strPictureName
            imgPicture.Picture = LoadPicture()
            MsgBox "Employee picture successfully deleted.", vbExclamation, Title
            
            Call sub_LOAD_TITLES(cboTitleOfCourtesy)
            Call sub_LOAD_SUPERVISORS(cboReportsTo)
        
            Call fn_DISABLE_CONTROLS
            Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
            fraEmployeeDetails.Enabled = False
            fraPicture.Enabled = False
            Call sub_LOAD_EMPLOYEES(lstEmployees)
            lstEmployees.Enabled = True
    End If
    
    
    
        
EXITPROCEDURE:
    Exit Sub
End Sub

Private Sub cmdSave_Click()

    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtFirstName, "Kindly enter the employee first name.") Then GoTo EXITPROCEDURE
    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtLastName, "Kindly enter the employee last name.") Then GoTo EXITPROCEDURE
    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtTitle, "Kindly enter his title(Position).") Then GoTo EXITPROCEDURE
    If Mdl_FUNCTIONS.fn_REQUIRE_COMBO_FIELD(cboTitleOfCourtesy, "Kindly enter the employee title of courtesy.") Then GoTo EXITPROCEDURE
    If Mdl_FUNCTIONS.fn_REQUIRE_DATE_OF_BIRTH(DTBirthDate, "The employee Birth Date can not be today or Tomorrow.") Then GoTo EXITPROCEDURE

    

    With cls_EMPLOYEES_Obj
        .EmployeeNo = Trim(cmdIdentification.Caption)
        .FirstName = Trim(txtFirstName.Text)
        .LastName = Trim(txtLastName.Text)
        .Title = Trim(txtTitle.Text)
        .TitleOfCourtesy = cboTitleOfCourtesy.ItemData(cboTitleOfCourtesy.ListIndex)
        .BirthDate = DTBirthDate.Value
        .HireDate = DTHireDate.Value
        .Address = Trim(txtAddress.Text)
        .City = Trim(txtCity.Text)
        .Region = Trim(txtRegion.Text)
        .PostalCode = Trim(txtPostalCode.Text)
        .Country = Trim(txtCountry.Text)
        .HomePhone = Trim(txtHomePhone.Text)
        .Extension = Trim(txtExtension.Text)
        .Notes = Trim(txtNotes.Text)
        If cboReportsTo.ListIndex = -1 Then
            .ReportsTo = 0
            Else
            .ReportsTo = cboReportsTo.ItemData(cboReportsTo.ListIndex)
        End If
        If ChkAlerte.Value = 0 Then
            .WorkingStatus = 0
            Else
                .WorkingStatus = 1
        End If

        If imgPicture.Picture Then
            
            strPictureName = Trim(cmdIdentification.Caption) & ".bmp"
    
            Call SavePicture(imgPicture, strPicturePath & strPictureName)
    
            .Photo = strPictureName
            
            Else
            
                .Photo = ""
                
        End If
    End With


    If blnEmployeeAdd = True And blnEmployeeModify = False Then
        cls_EMPLOYEES_Obj.fn_SAVE_EMPLOYEE_RECORDS
        MsgBox "Employee successfully saved.", vbExclamation, Title
        Else
            cls_EMPLOYEES_Obj.fn_UPDATE_EMPLOYEE_RECORDS (lngID)
            MsgBox "Employee successfully updated.", vbExclamation, Title
    End If

    Call sub_LOAD_TITLES(cboTitleOfCourtesy)
    Call sub_LOAD_SUPERVISORS(cboReportsTo)

    Call fn_DISABLE_CONTROLS
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    fraEmployeeDetails.Enabled = False
    fraPicture.Enabled = False
    Call sub_LOAD_EMPLOYEES(lstEmployees)
'    Call subLoadEmployeeDetails(lngID)
    lstEmployees.Enabled = True
    
    
    
EXITPROCEDURE:
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    blnEmployeeAdd = False
    blnEmployeeModify = True
    Call fn_ENABLE_CONTROLS
    Call Mdl_FUNCTIONS.fn_UNSET_CONTROL_COLOR(Me, "All")
    fraEmployeeDetails.Enabled = True
    fraPicture.Enabled = True
    lstEmployees.Enabled = False
    txtFirstName.SetFocus
End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_FORM_SIZE(Me)
    Move 0, 0
    Call fn_SET_CONTROL_COLOR(Me, "All")
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call fn_DISABLE_CONTROLS
    Call sub_LOAD_TITLES(cboTitleOfCourtesy)
    Call sub_LOAD_SUPERVISORS(cboReportsTo)
    Call sub_LOAD_EMPLOYEES(lstEmployees)

    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdAddNew, cmdAddNew.hwnd, "Add new employee.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdEdit, cmdEdit.hwnd, "Edit existing employee details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdDelete, cmdDelete.hwnd, "Delete existing employee details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSave, cmdSave.hwnd, "Save employee details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancel, cmdCancel.hwnd, "Cancel the transaction.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
'    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdPrint, cmdPrint.hwnd, "Print the list of all the employee", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdClose, cmdClose.hwnd, "Close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor

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

Private Sub sub_LOAD_TITLES(Optional cbo As ComboBox)

    Dim rec As New ADODB.Recordset
    Set rec = cls_REFERENCES_Obj.fn_LOADT_TITLES
    
    cbo.Clear
    
    Do While Not rec.EOF
        cbo.AddItem rec!TitleName
        cbo.ItemData(cbo.NewIndex) = rec!TitleID
        rec.MoveNext
    Loop

End Sub


Private Sub sub_LOAD_SUPERVISORS(Optional cbo As ComboBox)

    Dim rec As New ADODB.Recordset
    Set rec = cls_REFERENCES_Obj.fn_LOAD_SUPERVISORS(0)
    
    cbo.Clear
    
    Do While Not rec.EOF
        cbo.AddItem rec!FirstName & " " & rec!LastName
        cbo.ItemData(cbo.NewIndex) = rec!EmployeeID
        rec.MoveNext
    Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)
    blnEmployeeAdd = False
    blnEmployeeModify = False
End Sub

Private Sub lstEmployees_Click()
'    Call MdlFunctions.fnEmptyFields(Me)
    lngID = lstEmployees.ItemData(lstEmployees.ListIndex)
    Call subLoadEmployeeDetails(lstEmployees.ItemData(lstEmployees.ListIndex))
End Sub

Private Sub subLoadEmployeeDetails(lngEmployeeID As Long)
On Error GoTo errHandler
    Dim rec As New ADODB.Recordset
    
    Set rec = cls_EMPLOYEES_Obj.fn_LOAD_EMPLOYEES(lngEmployeeID)
    
    cmdIdentification.Caption = rec!EmployeeNo
    txtFirstName.Text = rec!FirstName & " "
    txtLastName.Text = rec!LastName & " "
    txtTitle.Text = rec!Title & " "
    cboTitleOfCourtesy.ListIndex = fn_GET_LIST_INDEX(cboTitleOfCourtesy, rec!TitleOfCourtesy)
    DTBirthDate.Value = rec!BirthDate & " "
    DTHireDate.Value = rec!HireDate & " "
    txtAddress.Text = rec!Address & " "
    txtCity.Text = rec!City & " "
    txtRegion.Text = rec!Region & " "
    txtPostalCode.Text = rec!PostalCode & " "
    txtCountry.Text = rec!Country & " "
    txtHomePhone.Text = rec!HomePhone & " "
    txtExtension.Text = rec!Extension & " "
    txtNotes.Text = rec!Notes & " "
    cboReportsTo.ListIndex = fn_GET_LIST_INDEX(cboReportsTo, rec!ReportsTo)
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
    MsgBox "Error Occurred while loading picture", vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, Me.Name, "cmdTelecharger")
    
    GoTo EXITPROCEDURE
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_ALPHABET_ONLY(KeyAscii)
End Sub

Private Sub txtHomePhone_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_ALPHABET_ONLY(KeyAscii)
End Sub

Private Sub txtPostalCode_Validate(Cancel As Boolean)
    txtPostalCode.Text = StrConv(txtPostalCode.Text, vbUpperCase)
End Sub

Public Function fn_DISABLE_CONTROLS()

    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdAddNew.Enabled = True

    cmdLoadPicture.Enabled = False
    cmdRemovePicture.Enabled = False

End Function

'**********Function to enable some buttons***********
Public Function fn_ENABLE_CONTROLS()

    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdAddNew.Enabled = False

    cmdLoadPicture.Enabled = True
    cmdRemovePicture.Enabled = True


End Function

