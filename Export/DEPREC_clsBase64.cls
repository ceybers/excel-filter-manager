VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DEPREC_clsBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type tBase64
    objXML As Variant
    objNode As Variant
    arrData() As Byte
End Type

Private this As tBase64

Function Encode(ByVal text As String) As String
    this.arrData = StrConv(text, vbFromUnicode)
    this.objNode.nodeTypedValue = this.arrData
    Encode = this.objNode.text
End Function

Function Decode(ByVal text As String) As String
    this.objNode.text = text
    this.arrData = this.objNode.nodeTypedValue
    Decode = StrConv(this.arrData, vbUnicode)
End Function

Private Sub Class_Initialize()
    Set this.objXML = CreateObject("MSXML2.DOMDocument")
    Set this.objNode = this.objXML.createElement("b64")
    this.objNode.DataType = "bin.base64"
End Sub

Private Sub Class_Terminate()
    Set this.objXML = Nothing
    Set this.objNode = Nothing
End Sub
