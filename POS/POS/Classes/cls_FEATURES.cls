VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_FEATURES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarFeatureID As Long 'local copy
Private mvarFeature As String 'local copy
Public Property Let Feature(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Feature = 5
    mvarFeature = vData
End Property


Public Property Get Feature() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Feature
    Feature = mvarFeature
End Property



Public Property Let FeatureID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FeatureID = 5
    mvarFeatureID = vData
End Property


Public Property Get FeatureID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FeatureID
    FeatureID = mvarFeatureID
End Property



