VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_ADD_PACKAGES 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PACKAGES"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin OCX.b8Container b8Container1 
      Height          =   4725
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   8334
      BorderColor     =   12735512
      BackColor       =   16185592
      Begin OCX.b8Container ContainerCatDetails 
         Height          =   3885
         Left            =   90
         TabIndex        =   6
         Top             =   90
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   6853
         BackColor       =   16185592
         Begin VB.TextBox txtTotalTax 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   500
            TabIndex        =   24
            Top             =   3000
            Width           =   3195
         End
         Begin VB.TextBox txtSellingPriceWithTax 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1920
            MaxLength       =   500
            TabIndex        =   3
            Top             =   1770
            Width           =   3195
         End
         Begin VB.TextBox txtSellingPriceWithoutTax 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   500
            TabIndex        =   4
            Top             =   3450
            Width           =   3195
         End
         Begin VB.CheckBox ChkNHIL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1920
            TabIndex        =   19
            Top             =   4650
            Width           =   315
         End
         Begin VB.CheckBox ChkVAT 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1920
            TabIndex        =   18
            Top             =   4230
            Width           =   225
         End
         Begin VB.TextBox txtQty 
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
            Left            =   1920
            MaxLength       =   100
            TabIndex        =   1
            Top             =   870
            Width           =   3195
         End
         Begin VB.TextBox txtSupplierPrice 
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
            Left            =   1920
            MaxLength       =   500
            TabIndex        =   2
            Top             =   1320
            Width           =   3195
         End
         Begin VB.TextBox txtVat 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   500
            TabIndex        =   13
            Top             =   2190
            Width           =   3195
         End
         Begin VB.TextBox txtNHIL 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   500
            TabIndex        =   12
            Top             =   2610
            Width           =   3195
         End
         Begin VB.ComboBox cboPackages 
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
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   420
            Width           =   1845
         End
         Begin lvButton.lvButtons_H cmdLoadPackages 
            Height          =   375
            Left            =   3810
            TabIndex        =   11
            Top             =   390
            Width           =   1335
            _ExtentX        =   2355
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
            Image           =   "frm_ADD_PACKAGES.frx":0000
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_ADD_PACKAGES.frx":0515
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Tax"
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
            TabIndex        =   25
            Top             =   3000
            Width           =   1365
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Selling Price + TAX"
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
            TabIndex        =   23
            Top             =   1770
            Width           =   1785
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Selling Price - TAX"
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
            TabIndex        =   22
            Top             =   3450
            Width           =   1635
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exempt VAT"
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
            Left            =   180
            TabIndex        =   21
            Top             =   4230
            Width           =   1050
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exempt NHIL"
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
            Left            =   180
            TabIndex        =   20
            Top             =   4650
            Width           =   1125
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Price"
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
            TabIndex        =   17
            Top             =   1320
            Width           =   1365
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
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
            TabIndex        =   16
            Top             =   870
            Width           =   1365
         End
         Begin VB.Label Label1 
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
            Height          =   345
            Left            =   180
            TabIndex        =   15
            Top             =   2190
            Width           =   525
         End
         Begin VB.Label Label4 
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
            Height          =   345
            Left            =   180
            TabIndex        =   14
            Top             =   2610
            Width           =   585
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Package"
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
            Left            =   210
            TabIndex        =   7
            Top             =   420
            Width           =   1155
         End
      End
      Begin OCX.b8Container b8Container6 
         Height          =   675
         Left            =   90
         TabIndex        =   8
         Top             =   3990
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   1191
         BackColor       =   16185592
         Begin lvButton.lvButtons_H cmdCancel 
            Height          =   495
            Left            =   3825
            TabIndex        =   9
            Top             =   90
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
            Image           =   "frm_ADD_PACKAGES.frx":082F
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_ADD_PACKAGES.frx":0D9E
         End
         Begin lvButton.lvButtons_H cmdSave 
            Height          =   495
            Left            =   2460
            TabIndex        =   10
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
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
            Image           =   "frm_ADD_PACKAGES.frx":10B8
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_ADD_PACKAGES.frx":1318
         End
      End
   End
End
Attribute VB_Name = "frm_ADD_PACKAGES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngPackageID As Long

Private Sub cboPackages_Click()
    lngPackageID = cboPackages.ItemData(cboPackages.ListIndex)
End Sub

Private Sub ChkNHIL_Click()
    If ChkNHIL.Value = 0 Then
        txtNHIL.Text = Val(txtSellingPriceWithoutTax.Text) * lngNHIL / 100
        txtSellingPriceWithTax.Text = Val(txtSellingPriceWithoutTax.Text) + Val(txtNHIL.Text) + Val(txtVat.Text)
        ElseIf ChkNHIL.Value = 1 Then
            txtNHIL.Text = 0
            txtSellingPriceWithTax.Text = Val(txtSellingPriceWithoutTax.Text) - Val(txtNHIL.Text) + Val(txtVat.Text)
    End If
End Sub

Private Sub ChkVAT_Click()

    If ChkVAT.Value = 0 Then
        txtVat.Text = Val(txtSellingPriceWithoutTax.Text) * lngVAT / 100
        txtSellingPriceWithTax.Text = Val(txtSellingPriceWithoutTax.Text) + Val(txtVat.Text) + Val(txtNHIL.Text)
        ElseIf ChkVAT.Value = 1 Then
            txtVat.Text = 0
           txtSellingPriceWithTax.Text = Val(txtSellingPriceWithoutTax.Text) - Val(txtVat.Text) + Val(txtNHIL.Text)
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLoadPackages_Click()
    blnAddPackage = True
    frm_PACKAGES.Show 1
End Sub

Private Sub cmdSave_Click()

    With frm_PRODUCTS
        Set lstItem = .lvw.ListItems.Add(, , lngPackageID)
            lstItem.ListSubItems.Add , , cboPackages.Text
            lstItem.ListSubItems.Add , , Trim(txtQty.Text)
            lstItem.ListSubItems.Add , , Trim(txtSupplierPrice.Text)
            lstItem.ListSubItems.Add , , Trim(txtSellingPriceWithoutTax.Text)
            lstItem.ListSubItems.Add , , Trim(txtVat.Text)
            lstItem.ListSubItems.Add , , Trim(txtNHIL.Text)
            lstItem.ListSubItems.Add , , Trim(txtSellingPriceWithTax.Text)
        
    End With
    
    blnEditPriceControl = False
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    Call Mdl_FUNCTIONS.sub_CENTER_FORM(Me)
    Call sub_LOAD_PACKAGES
    
    ChkVAT.Value = 1
    ChkNHIL.Value = 1


    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSave, cmdSave.hwnd, "Save package details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancel, cmdCancel.hwnd, "Cancel and close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdLoadPackages, cmdLoadPackages.hwnd, "Add new package to the existing ones.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If blnEditPriceControl = True Then
        Call cmdSave_Click
    End If
End Sub

Public Sub sub_LOAD_PACKAGES()
    Dim rec As New ADODB.Recordset
    
    Set rec = cls_REFERENCES_Obj.fn_LOAD_PACKAGES(0)
    
    cboPackages.Clear
    
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            cboPackages.AddItem rec!PackageName
            cboPackages.ItemData(cboPackages.NewIndex) = rec!PackageID
            rec.MoveNext
        Loop
    End If
    
End Sub



Private Sub txtSellingPriceWithoutTax_Change()
'    txtVat.Text = Val(txtSellingPriceWithoutTax.Text) * lngVAT / 100
'    txtNHIL.Text = Val(txtSellingPriceWithoutTax.Text) * lngNHIL / 100
End Sub

Private Sub txtSellingPriceWithoutTax_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtSellingPriceWithoutTax_LostFocus()
'    txtSellingPriceWithoutTax.Text = Format(IIf(txtSellingPriceWithoutTax.Text = "", 0, txtSellingPriceWithoutTax.Text), "#,##0.00")
'    txtSellingPriceWithoutTax.Alignment = 1
End Sub

Private Sub txtSellingPriceWithTax_Change()
    txtTotalTax.Text = Val(txtSellingPriceWithTax.Text) * Val(lngVAT + lngNHIL) / 100
    txtSellingPriceWithoutTax.Text = Val(txtSellingPriceWithTax.Text) - Val(txtTotalTax.Text)

    txtVat.Text = Val(txtTotalTax.Text) / Val(lngVAT + lngNHIL) * lngVAT
    txtNHIL.Text = Val(txtTotalTax.Text) / Val(lngVAT + lngNHIL) * lngNHIL
End Sub

Private Sub txtSupplierPrice_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtSupplierPrice_LostFocus()
'    txtSupplierPrice.Text = Format(IIf(txtSupplierPrice.Text = "", 0, txtSupplierPrice.Text), "#,##0.00")
'    txtSupplierPrice.Alignment = 1
End Sub
