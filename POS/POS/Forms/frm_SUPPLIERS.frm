VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_SUPPLIERS 
   Caption         =   "SUPPLIERS"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   12180
   Begin OCX.b8Container b8Container2 
      Height          =   4005
      Left            =   0
      TabIndex        =   11
      Top             =   4560
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   7064
      BackColor       =   16185592
      Begin VB.Frame fraAllProducts 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3765
         Left            =   120
         TabIndex        =   15
         Top             =   150
         Width           =   4875
         Begin MSComctlLib.ListView lvwProducts 
            Height          =   3405
            Left            =   90
            TabIndex        =   16
            Top             =   330
            Width           =   4665
            _ExtentX        =   8229
            _ExtentY        =   6006
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Product Name"
               Object.Width           =   8204
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
         End
         Begin OCX.b8SideTab b8SideTab1 
            Height          =   375
            Left            =   90
            TabIndex        =   17
            Top             =   0
            Width           =   4665
            _ExtentX        =   8229
            _ExtentY        =   661
            Caption         =   "ALL PRODUCTS"
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
      Begin VB.Frame fraSupplierProd 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3795
         Left            =   6990
         TabIndex        =   12
         Top             =   120
         Width           =   4875
         Begin MSComctlLib.ListView lvwSupplierProducts 
            Height          =   3405
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   4665
            _ExtentX        =   8229
            _ExtentY        =   6006
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Product Name"
               Object.Width           =   8203
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
         End
         Begin OCX.b8SideTab b8SideTab3 
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   30
            Width           =   4665
            _ExtentX        =   8229
            _ExtentY        =   661
            Caption         =   "SUPPLIER PRODUCTS"
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
      Begin lvButton.lvButtons_H cmdRemove 
         Height          =   465
         Left            =   5130
         TabIndex        =   18
         Top             =   2190
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   820
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
         mIcon           =   "frm_SUPPLIERS.frx":0000
      End
      Begin lvButton.lvButtons_H cmdAdd 
         Height          =   465
         Left            =   5130
         TabIndex        =   19
         Top             =   1470
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   820
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
         mIcon           =   "frm_SUPPLIERS.frx":031A
      End
   End
   Begin OCX.b8Container b8Container5 
      Height          =   4575
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   8070
      BackColor       =   16185592
      Begin OCX.b8Container b8Container6 
         Height          =   675
         Left            =   60
         TabIndex        =   21
         Top             =   3870
         Width           =   11865
         _ExtentX        =   20929
         _ExtentY        =   1191
         BackColor       =   16185592
         Begin lvButton.lvButtons_H cmdAddNew 
            Height          =   495
            Left            =   180
            TabIndex        =   22
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
            Image           =   "frm_SUPPLIERS.frx":0634
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_SUPPLIERS.frx":0B49
         End
         Begin lvButton.lvButtons_H cmdCancel 
            Height          =   495
            Left            =   8172
            TabIndex        =   23
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
            Image           =   "frm_SUPPLIERS.frx":0E63
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_SUPPLIERS.frx":13D2
         End
         Begin lvButton.lvButtons_H cmdEdit 
            Height          =   495
            Left            =   2178
            TabIndex        =   24
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
            Image           =   "frm_SUPPLIERS.frx":16EC
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_SUPPLIERS.frx":1A69
         End
         Begin lvButton.lvButtons_H cmdClose 
            Height          =   495
            Left            =   10170
            TabIndex        =   25
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
            Image           =   "frm_SUPPLIERS.frx":1D83
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_SUPPLIERS.frx":2313
         End
         Begin lvButton.lvButtons_H cmdSave 
            Height          =   495
            Left            =   6174
            TabIndex        =   26
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
            Image           =   "frm_SUPPLIERS.frx":262D
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_SUPPLIERS.frx":288D
         End
         Begin lvButton.lvButtons_H cmdDelete 
            Height          =   495
            Left            =   4176
            TabIndex        =   27
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
            Image           =   "frm_SUPPLIERS.frx":2BA7
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_SUPPLIERS.frx":30BA
         End
      End
      Begin OCX.b8Container b8Container4 
         Height          =   3765
         Left            =   4050
         TabIndex        =   28
         Top             =   90
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   6641
         BackColor       =   16185592
         Begin VB.Frame fraSupplierDetails 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   3645
            Left            =   90
            TabIndex        =   29
            Top             =   60
            Width           =   7545
            Begin VB.TextBox txtHomePage 
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
               Left            =   1740
               MaxLength       =   100
               TabIndex        =   10
               Top             =   3300
               Width           =   5805
            End
            Begin VB.TextBox txtPhone 
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
               Left            =   1740
               MaxLength       =   50
               TabIndex        =   8
               Top             =   2580
               Width           =   5805
            End
            Begin VB.TextBox txtFax 
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
               Left            =   1740
               MaxLength       =   50
               TabIndex        =   9
               Top             =   2940
               Width           =   5805
            End
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
               Left            =   4410
               MaxLength       =   100
               TabIndex        =   2
               Top             =   420
               Width           =   3135
            End
            Begin VB.ComboBox cboContactTitle 
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
               Left            =   1740
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   420
               Width           =   1275
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
               Left            =   1740
               MaxLength       =   50
               TabIndex        =   7
               Top             =   2220
               Width           =   5805
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
               Left            =   1740
               MaxLength       =   50
               TabIndex        =   6
               Top             =   1860
               Width           =   5805
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
               Left            =   1740
               MaxLength       =   100
               TabIndex        =   5
               Top             =   1500
               Width           =   5805
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
               Left            =   1740
               MaxLength       =   100
               TabIndex        =   4
               Top             =   1140
               Width           =   5805
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
               Left            =   1740
               MaxLength       =   200
               TabIndex        =   3
               Top             =   780
               Width           =   5805
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
               Left            =   1740
               MaxLength       =   100
               TabIndex        =   0
               Top             =   60
               Width           =   5805
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Web Site"
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
               TabIndex        =   40
               Top             =   3300
               Width           =   1125
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Phone No"
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
               TabIndex        =   39
               Top             =   2580
               Width           =   1215
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Fax"
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
               TabIndex        =   38
               Top             =   2940
               Width           =   1125
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Title"
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
               Top             =   420
               Width           =   555
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
               Left            =   150
               TabIndex        =   36
               Top             =   2220
               Width           =   1125
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "B.O.Pox"
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
               TabIndex        =   35
               Top             =   1860
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
               Left            =   150
               TabIndex        =   34
               Top             =   1500
               Width           =   1035
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Town"
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
               TabIndex        =   33
               Top             =   1140
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
               Left            =   150
               TabIndex        =   32
               Top             =   780
               Width           =   1065
            End
            Begin VB.Label Label3 
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
               Left            =   3150
               TabIndex        =   31
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label2 
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
               Left            =   150
               TabIndex        =   30
               Top             =   60
               Width           =   2265
            End
         End
      End
      Begin OCX.b8Container b8Container3 
         Height          =   3765
         Left            =   60
         TabIndex        =   41
         Top             =   90
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   6641
         BackColor       =   16185592
         Begin VB.ListBox lstSuppliers 
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
            Height          =   3150
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   42
            Top             =   480
            Width           =   3765
         End
         Begin OCX.b8SideTab b8SideTab2 
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   150
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   661
            Caption         =   "SUPPLIERS"
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
   End
End
Attribute VB_Name = "frm_SUPPLIERS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lngID As Long
Public lngProID As Long

Private Sub sub_LOAD_SUPPLIER_PRODUCT(lngSupplierID As Long)

    Dim rec As New ADODB.Recordset
    Set rec = cls_PRODUCT_Obj.fn_LOAD_SUPPLIER_PRODUCT(lngSupplierID)
    
    lvwSupplierProducts.ListItems.Clear
    
    Do While Not rec.EOF
        Set lvwItem = lvwSupplierProducts.ListItems.Add(, , rec!ProductName)
        lvwItem.ListSubItems.Add , , rec!ProductID
        rec.MoveNext
    Loop

End Sub


Private Sub sub_LOAD_PRODUCT(Optional lngSupplierID As Long)

    Dim rec As New ADODB.Recordset
    Set rec = cls_PRODUCT_Obj.fN_LOAD_PRODUCTS(lngSupplierID)
    
    lvwProducts.ListItems.Clear
    
    Do While Not rec.EOF
        Set lvwItem = lvwProducts.ListItems.Add(, , rec!ProductName)
        lvwItem.ListSubItems.Add , , rec!ProductID
        rec.MoveNext
    Loop

End Sub

Private Sub sub_LOAD_SUPPLIERS(Optional lst As ListBox)

    Dim rec As New ADODB.Recordset
    Set rec = cls_SUPPLIER_Obj.fn_LOAD_SUPPLIERS(0)
    
    lst.Clear
    
    Do While Not rec.EOF
        lst.AddItem rec!CompanyName
        lst.ItemData(lst.NewIndex) = rec!SupplierID
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

Private Sub cmdAdd_Click()
On Error GoTo errHandler
    Dim ctr As Integer
    Dim ctrLvwProducts As Integer
    Dim ctrLvwSupplierProducts As Integer
    
    If lngID = 0 Then
        MsgBox "Please select the supplier", vbExclamation, Title
        GoTo EXITPROCEDURE
    End If
    
    
    For ctrLvwProducts = 1 To lvwProducts.ListItems.Count
        If lvwProducts.ListItems(ctrLvwProducts).Checked = True Then
            Call cls_PRODUCT_Obj.fn_CHECK_SUPPLIER_PRODUCT(lngID, lvwProducts.ListItems(ctrLvwProducts).ListSubItems(1).Text)
            If blnCheckSupplierProductExist = False Then
                Call cls_PRODUCT_Obj.fn_UPDATE_SUPPLIER_PRODUCT(lngID, lvwProducts.ListItems(ctrLvwProducts).ListSubItems(1).Text)
                lvwProducts.ListItems(ctrLvwProducts).Checked = False
            End If
        End If
    Next
    
    
    For ctr = 1 To lvwProducts.ListItems.Count
        lvwProducts.ListItems(ctr).Checked = False
    Next
    
    Call lstSuppliers_Click
    
EXITPROCEDURE:
    Exit Sub

errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
    
End Sub


Private Sub cmdaddNew_Click()
    blnAddSupplier = True
    blnEditSupplier = False
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call Mdl_FUNCTIONS.fn_ENABLE_CONTROLS(Me)
    Call fn_UNSET_CONTROL_COLOR(Me, "All")
    fraSupplierDetails.Enabled = True
    fraAllProducts.Enabled = True
    fraSupplierProd.Enabled = True
    lstSuppliers.Enabled = False
    txtCompanyName.SetFocus
End Sub

Private Sub cmdDelete_Click()
On Error GoTo errHandler
    If MsgBox("Are you sure you want to Delete Supplier " & Trim(txtCompanyName.Text) & "?", vbQuestion + vbYesNo, Title) = vbNo Then Exit Sub

    With cls_SUPPLIER_Obj
        .fn_CHECK_SUPPLIER_IN_ORDERS (lngID)
        If blnSupplierExist = True Then
            MsgBox "Supplier Can Not Be Deleted.", vbInformation, Title
            Exit Sub
            Else
                Call .fn_DELETE_SUPPLIER(lngID)
                Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
                Call sub_LOAD_SUPPLIERS(lstSuppliers)
        End If
    End With

EXITPROCEDURE:
    Exit Sub

errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub

Private Sub cmdSave_Click()
On Error GoTo errHandler

    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtCompanyName, "Please enter the company name!") Then GoTo EXITPROCEDURE
    If Mdl_FUNCTIONS.fn_REQUIRE_COMBO_FIELD(cboContactTitle, "Please select the contact title!") Then GoTo EXITPROCEDURE
    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtContactName, "Please enter the contact name!") Then GoTo EXITPROCEDURE

    
    With cls_SUPPLIER_Obj
        .SupplierID = .fn_AUTOGEN
        .CompanyName = Trim(txtCompanyName.Text)
        .ContactTitle = cboContactTitle.ItemData(cboContactTitle.ListIndex)
        .ContactName = Trim(txtContactName.Text)
        .Address = Trim(txtAddress.Text)
        .City = Trim(txtCity.Text)
        .Region = Trim(txtRegion.Text)
        .PostalCode = Trim(txtPostalCode.Text)
        .Country = Trim(txtCountry.Text)
        .Phone = Trim(txtPhone.Text)
        .Fax = Trim(TxtFax.Text)
        .HomePage = Trim(txtHomePage.Text)
    
        If blnAddSupplier = True And blnEditSupplier = False Then
            .fn_SAVE_SUPPLIERS_RECORDS
            MsgBox "Suppliers details saved successfully", vbExclamation, Title
            Else
                .fn_UPDATE_SUPPLIERS_RECORDS (lngID)
                MsgBox "Suppliers details saved successfully", vbExclamation, Title
        End If
    End With
    
    Call fn_DISABLE_CONTROLS(Me)
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    fraSupplierDetails.Enabled = False
    fraAllProducts.Enabled = False
    fraSupplierProd.Enabled = False
    Call sub_LOAD_SUPPLIERS(lstSuppliers)
    Call sub_LOAD_SUPPLIERS_DETAILS(lngID)
    lstSuppliers.Enabled = True
    
EXITPROCEDURE:
    Exit Sub

errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub



Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Call fn_DISABLE_CONTROLS(Me)
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    fraSupplierDetails.Enabled = False
    fraAllProducts.Enabled = False
    fraSupplierProd.Enabled = False
    lstSuppliers.Enabled = True
    Call lstSuppliers_Click
End Sub

Private Sub cmdEdit_Click()
    blnAddSupplier = False
    blnEditSupplier = True
    Call Mdl_FUNCTIONS.fn_ENABLE_CONTROLS(Me)
    Call Mdl_FUNCTIONS.fn_UNSET_CONTROL_COLOR(Me, "All")
    fraSupplierDetails.Enabled = True
    fraAllProducts.Enabled = True
    fraSupplierProd.Enabled = True
    lstSuppliers.Enabled = False
    txtCompanyName.SetFocus
End Sub



Private Sub cmdRemove_Click()
On Error GoTo errHandler

    Dim ctr As Integer
    
    For ctr = 1 To lvwSupplierProducts.ListItems.Count
        If lvwSupplierProducts.ListItems(ctr).Checked = True Then
            Call cls_PRODUCT_Obj.fn_DELETE_SUPPLIER_PRODUCT(lngID, lvwSupplierProducts.ListItems(ctr).ListSubItems(1).Text)
'            lvwSupplierProducts.ListItems.Remove lvwSupplierProducts.ListItems(ctr).Index
        End If
    Next
    
    Call lstSuppliers_Click
    
EXITPROCEDURE:
    Exit Sub

errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
    
End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_FORM_SIZE(Me)
    Move 0, 0
    Call fn_SET_CONTROL_COLOR(Me, "All")
    Call fn_DISABLE_CONTROLS(Me)
    Call sub_LOAD_TITLES(cboContactTitle)
    Call sub_LOAD_SUPPLIERS(lstSuppliers)
    Call sub_LOAD_PRODUCT
    fraAllProducts.Enabled = False
    fraSupplierProd.Enabled = False
    
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdAddNew, cmdAddNew.hwnd, "Add new supplier details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdEdit, cmdEdit.hwnd, "Edit existing supplier details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdDelete, cmdDelete.hwnd, "Delete existing supplier details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSave, cmdSave.hwnd, "Save supplier details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancel, cmdCancel.hwnd, "Cancel the transaction.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdClose, cmdClose.hwnd, "Close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdAdd, cmdAdd.hwnd, "Add product to supplier list.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdRemove, cmdRemove.hwnd, "Remove product from supplier list.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blnAddSupplier = False
    blnEditSupplier = False
    blnCheckSupplierProductExist = False
End Sub

Private Sub lstSuppliers_Click()
    If lstSuppliers.ListIndex = -1 Then Exit Sub
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    lngID = lstSuppliers.ItemData(lstSuppliers.ListIndex)
    Call sub_LOAD_SUPPLIERS_DETAILS(lstSuppliers.ItemData(lstSuppliers.ListIndex))
    Call sub_LOAD_SUPPLIER_PRODUCT(lstSuppliers.ItemData(lstSuppliers.ListIndex))
    b8SideTab3.Caption = "Products Supplied By " & lstSuppliers.Text
End Sub

Private Sub sub_LOAD_SUPPLIERS_DETAILS(lngSupplierID As Long)

    Dim rec As New ADODB.Recordset
    
    Set rec = cls_SUPPLIER_Obj.fn_LOAD_SUPPLIERS(lngSupplierID)
    
    
    With ClsSupplierObject
    
        txtCompanyName.Text = rec!CompanyName
        txtContactName.Text = rec!ContactName
        cboContactTitle.ListIndex = Mdl_FUNCTIONS.fn_GET_LIST_INDEX(cboContactTitle, rec!ContactTitle)
        txtAddress.Text = rec!Address
        txtCity.Text = rec!City
        txtRegion.Text = rec!Region
        txtPostalCode.Text = rec!PostalCode
        txtCountry.Text = rec!Country
        txtPhone.Text = rec!Phone
        TxtFax.Text = rec!Fax
        txtHomePage.Text = rec!HomePage
        
    End With
    
End Sub

Private Sub txtCity_Validate(Cancel As Boolean)
    txtCity.Text = StrConv(txtCity.Text, vbProperCase)
End Sub

Private Sub txtContactName_Validate(Cancel As Boolean)
    txtContactName.Text = StrConv(txtContactName.Text, vbProperCase)
End Sub

Private Sub txtCountry_Validate(Cancel As Boolean)
    txtCountry.Text = StrConv(txtCountry.Text, vbProperCase)
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtPostalCode_Validate(Cancel As Boolean)
    txtPostalCode.Text = StrConv(txtPostalCode.Text, vbUpperCase)
End Sub

Private Sub txtRegion_Validate(Cancel As Boolean)
    txtRegion.Text = StrConv(txtRegion.Text, vbProperCase)
End Sub
