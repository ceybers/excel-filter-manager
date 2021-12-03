Attribute VB_Name = "modBase64"
Option Explicit

Private Function CachedBase64Node() As Object
    Static objXML As Object
    Static objNode As Object
    Static Ready As Boolean
    
    If Ready = False Then
        Set objXML = CreateObject("MSXML2.DOMDocument")
        Set objNode = objXML.createElement("b64")
        objNode.DataType = "bin.base64"
        Ready = True
    Else
        'Debug.Print "Do not need to recache Static64"
    End If
    
    Set CachedBase64Node = objNode
End Function

Function StringToBase64(ByVal text As String) As String
    Dim arrData() As Byte
    Dim cachedNode As Object
    
    Set cachedNode = CachedBase64Node()
    arrData = StrConv(text, vbFromUnicode)
    cachedNode.nodeTypedValue = arrData
    StringToBase64 = cachedNode.text
End Function

Function Base64toString(ByVal text As String) As String
    Dim arrData() As Byte
    Dim cachedNode As Object
    
    Set cachedNode = CachedBase64Node()
    
    cachedNode.text = text
    arrData = cachedNode.nodeTypedValue
    Base64toString = StrConv(arrData, vbUnicode)
End Function
