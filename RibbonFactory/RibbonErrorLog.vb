Imports System.Xml.Schema

Public NotInheritable Class RibbonErrorLog
    Private Readonly _errors As ICollection(Of String)

    Public Sub New(ribbonX As String, ParamArray xmlSchemas As XmlSchema())
        NullGuard.NotNull(xmlSchemas, NameOf(xmlSchemas))
        _errors = New List(Of String)
        ValidateRibbon(ribbonX, xmlSchemas)
    End Sub

    Public ReadOnly Property None As Boolean
        Get
            Return Not _errors.Any()
        End Get
    End Property

    Public ReadOnly Property Errors as IEnumerable(Of String)
        Get
            Return _errors
        End Get
    End Property

    Private Sub ValidateRibbon(ribbonX As String, ParamArray xmlSchemas As XmlSchema())
        Dim xsd As XmlSchemaSet = New XmlSchemaSet()

        Array.ForEach(xmlSchemas, Sub(schema) xsd.Add(schema))
        
        XDocument.Parse(ribbonX).Validate(xsd, AddressOf LogXmlError)
    End Sub

    Private Sub LogXmlError(sender As Object, e As ValidationEventArgs)
        _errors.Add($"{DirectCast(sender, XAttribute).Parent.Name.LocalName} {e.Severity.ToString().ToUpper}: {e.Message}")
    End Sub
End Class