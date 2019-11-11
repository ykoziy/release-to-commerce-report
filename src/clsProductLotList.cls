Attribute VB_Name = "clsProductLotList"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' Class Module: clsProductLotList
' This class represent a single item number and list of all the
' lot numbers relatd to it

Private strItemNumber As String
Private strLotNumber As String
Private strDescription As String

'Private dictLots As Scripting.dictionary
Private arrList As ArrayList

Private intTotalQuantity As Long
Private dblUnitPrice As Double

Sub InitProductLotList(itemNumber As String, desc As String)
    intTotalQuantity = 0
    dblUnitPrice = 0
    strItemNumber = itemNumber
    strDescription = desc
    'Set dictLots = New Scripting.dictionary
    Set arrList = New ArrayList
End Sub

Sub InsertLot(lot As String, LotStatus As Variant, Location As Variant, locQty _
As Variant, onHand As Variant, UnitPrice As Variant, WrkOrder As Variant)
    Dim productLot As New clsProductLotInformation
    Call productLot.InitLotInformation(lot, LotStatus, Location, locQty, onHand, WrkOrder)
    dblUnitPrice = CDbl(UnitPrice)
    'If Not dictLots.Exists(lot) Then
    '    dictLots.Add lot, productLot
    'End If
    arrList.Add productLot
End Sub

Sub AddQuantity(qty As Long)
    intTotalQuantity = intTotalQuantity + qty
End Sub

Public Property Get TotalQuantity() As Long
    TotalQuantity = intTotalQuantity
End Property

Public Property Get Description() As String
    Description = strDescription
End Property

Public Property Get itemNumber() As String
    itemNumber = strItemNumber
End Property


Public Property Get CountIt() As String
    CountIt = dictLots.Count
End Property

Public Property Get LotList() As Scripting.dictionary
    Set LotList = dictLots
End Property

Public Property Get LotListArray() As ArrayList
    Set LotListArray = arrList
End Property

Public Property Get UnitPrice() As Double
    UnitPrice = dblUnitPrice
End Property
