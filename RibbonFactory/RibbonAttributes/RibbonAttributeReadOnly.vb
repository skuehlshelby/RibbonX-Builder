Imports RibbonFactory.Utilities

Namespace RibbonAttributes
    Friend Class RibbonAttributeReadOnly(Of T)
        Inherits RibbonAttribute

        Protected Value As T

        Public Sub New(value As T, name As AttributeName, category As AttributeCategory)
            MyBase.New(name, category)
            Me.Value = value
        End Sub

        Public Function GetValue() As T
            Return Value
        End Function

        Public Overrides ReadOnly Property Xml As String
            Get
                If TypeOf Value Is Boolean Then
                    Return String.Format(XmlTemplate, Name.CamelCase(), Value.CamelCase())
                ElseIf Value.ToString() = String.Empty
                    Return String.Empty
                Else 
                    Return String.Format(XmlTemplate, Name.CamelCase(), Value)
                End If
            End Get
        End Property
    End Class
End NameSpace