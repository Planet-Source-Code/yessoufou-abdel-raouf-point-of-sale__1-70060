Attribute VB_Name = "MdlGlobalVariables"
Public curPosition As Byte
Public Const msgToScroll As String = "AbdelSoft: Let's Think Of The Way Forward"
Public file As New FileSystemObject
Public txtStream As TextStream
Public db As New Cls_DATABASE
Public XNode As Node
Public Const Title As String = "AbdelSoft"
Public ctl As Control

Public cls_DATABASE_Obj As New Cls_DATABASE
Public cls_USER_Obj As New Cls_USER

Public strPicturePath As String

Public lvwItem As ListItem

Type RegistrySettingsType
    Server As String
    Dababase As String
    UserName As String
    Password As String
End Type

Public registrySettings As RegistrySettingsType

