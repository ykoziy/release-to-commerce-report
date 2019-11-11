Attribute VB_Name = "clsProductLotInformation"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' Class Module: clsProductLotInformation
' This class represent an item with specific lot and related information

Private strLotNumber As String
Private strLotStatus As String
Private strLocation As String
Private lngLocationQuantity As String
Private dtOnHandDate As Date
Private strWorkOrder As String

Sub InitLotInformation(LotNumber As Variant, LotStatus As Variant, Location As _
Variant, locQty As Variant, onHand As Variant, WrkOrder As Variant)
    strLotNumber = CStr(LotNumber)
    strLotStatus = CStr(LotStatus)
    strLocation = CStr(Location)
    lngLocationQuantity = CLng(locQty)
    dtOnHandDate = CDate(onHand)
    strWorkOrder = CStr(WrkOrder)
End Sub

Public Property Get LotNumber() As String
    LotNumber = strLotNumber
End Property

Public Property Get LotStatus() As String
    LotStatus = strLotStatus
End Property

Public Property Get Location() As String
    Location = strLocation
End Property

' WIP ------- !!!!!!!!!!!!!!!
Public Property Get WrkOrder() As String
    WrkOrder = strWorkOrder
End Property

Public Property Get LocationQuantity() As Long
    LocationQuantity = lngLocationQuantity
End Property

Public Property Get OnHandDate() As Date
    OnHandDate = dtOnHandDate
End Property
