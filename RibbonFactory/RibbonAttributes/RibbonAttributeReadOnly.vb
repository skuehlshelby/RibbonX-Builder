Namespace RibbonAttributes
    Friend Class RibbonAttributeReadOnly(Of T)
        Inherits RibbonAttribute

        Protected Value As T

        Public Sub New(value As T, name As AttributeName, ParamArray categoryMembers As AttributeName())
            MyBase.New(name, categoryMembers)
            Me.Value = value
        End Sub

        Public Function GetValue() As T
            Return Value
        End Function

        Public Overrides ReadOnly Property Xml As String
            Get
                Return String.Format(XmlTemplate, Name.CamelCase(), Value)
            End Get
        End Property
    End Class
End NameSpace