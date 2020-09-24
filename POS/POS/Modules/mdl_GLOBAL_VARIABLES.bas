Attribute VB_Name = "mdl_GLOBAL_VARIABLES"
Public curPosition As Byte
'Public Const msgToScroll As String = "XSoft: Let's Think Of The Way Forward"
Public Const msgToScroll As String = "POS: Point Of Sale"
Public WorkStationName As String
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public file As New FileSystemObject
Public txtStream As TextStream
Public db As New Cls_DATABASE
Public XNode As Node
Public Const Title As String = "XSoft"
Public ctl As Control
Public lstItem As ListItem

Type SystemDataType

    DB_ServerName As String
    DB_Database As String
    DB_UserName As String
    DB_Password As String
    
End Type

Public SystemData As SystemDataType

Public cls_DATABASE_Obj As New Cls_DATABASE
Public cls_USER_Obj As New Cls_USER
Public cls_USERS_ACCESS_LOG_Obj As New cls_USERS_ACCESS_LOG
Public cls_CATEGORY_Obj As New cls_CATEGORIES
Public cls_PRODUCT_Obj As New cls_PRODUCTS
Public cls_CUSTOMER_Obj As New cls_CUSTOMERS
Public cls_SUPPLIER_Obj As New cls_SUPPLIERS
Public cls_ORDER_Obj As New cls_ORDERS
Public cls_SALES_Obj As New cls_SALES
Public cls_DELIVERY_Obj As New cls_DELIVERY
Public cls_REFERENCES_Obj As New cls_REFERENCES
Public cls_COMPANY_INFO_Obj As New cls_COMPANY_INFO
Public cls_PRODUCT_PACKAGE_Obj As New cls_PRODUCT_PACKAGE
Public cls_EXPENDITURES_Obj As New cls_EXPENDITURES
Public cls_RECEIPTS_Obj As New cls_RECEIPTS
Public cls_CUSTOMERS_ORDERS_Obj As New cls_CUSTOMERS_ORDERS
Public cls_BANK_Obj As New cls_BANKS
Public cls_ACCOUNT_TYPE_Obj As New cls_ACCOUNT_TYPE
Public cls_CASH_Obj As New cls_CASH
Public cls_CHEQUE_Obj As New cls_CHEQUE
Public cls_CHEQUES_Obj As New cls_CHEQUES
Public cls_BANK_DEPOSIT_Obj As New cls_BANK_DEPOSIT
Public cls_CASH_CHEQUE_BANK_DEPOSIT_Obj As New cls_CASH_CHEQUE_BANK
Public cls_EMPLOYEES_Obj  As New Cls_EMPLOYEES
Public cls_SALARIES_Obj  As New cls_SALARIES
Public cls_LEAVES_Obj  As New cls_LEAVES
Public cls_SALARY_PERIOD_Obj  As New cls_SALARY_PERIOD
Public cls_BANK_TRANSACTION_Obj As New cls_BANK_TRANSACTION

Public lngVAT As Double
Public lngNHIL As Double
Public strCompanyName As String
Public strAddress As String
Public strEMail As String
Public strPhoneNo As String
Public strFax As String
Public strLocation As String
Public strVatNo As String

Public lngCurrentUserID As Long
Public lngSelectedCustomerID As Long
Public lngAdminID As Long
Public strUserName As String
Public strPassword As String
Public strFullName As String

Public blnCheckCategoryProductExist As Boolean
Public blnCheckSupplierProductExist As Boolean

Public blnEditCategory As Boolean
Public blnAddCategory As Boolean
Public blnEditPriceControl As Boolean

Public blnEditProduct As Boolean
Public blnAddProduct As Boolean

Public blnEditSupplier As Boolean
Public blnAddSupplier As Boolean

Public blnProductPackage As Boolean
Public blnAddPackage As Boolean
Public blnCategory As Boolean

Public blnCheckIfSalaryPaid As Boolean
Public blnCustomerExist As Boolean
Public blnSupplierExist As Boolean
Public blnCategoryExist As Boolean
Public blnCustomerDeposit As Boolean
Public blnNewDeposit As Boolean
Public blnNewWithdrawal As Boolean
Public blnNewBankCharges As Boolean

Public strPicturePath As String

Public lvwItem As ListItem


Public dblTotalSales As Double
Public blnPending As Boolean
Public blnCancelPayment As Boolean
Public blnPendingSales As Boolean
Public Const tooltipBackColor As Variant = vbWhite
Public Const tooltipForeColor As Variant = vbBlue
