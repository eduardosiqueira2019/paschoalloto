Attribute VB_Name = "cep"
Function ConsultaCep(ValorCep As String, TipoCampo As String)

    Dim XmlDoc As DOMDocument
    Dim XmlNode As IXMLDOMNode
    Dim XmlNodes As IXMLDOMNodeList
    
    Set XmlDoc = New DOMDocument
    
    XmlDoc.async = False
    XmlDoc.Load ("https://viacep.com.br/ws/" + ValorCep + "/xml/")
    
    Set XmlNodes = XmlDoc.selectNodes("/xmlcep/" + TipoCampo)
    
    For Each XmlNode In XmlNodes
        ConsultaCep = XmlNode.Text
    Next

End Function
