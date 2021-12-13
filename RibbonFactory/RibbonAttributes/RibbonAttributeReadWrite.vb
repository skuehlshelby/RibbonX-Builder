Imports RibbonFactory.Utilities

Namespace RibbonAttributes
    Friend NotInheritable Class RibbonAttributeReadWrite(Of T)
        Inherits RibbonAttributeReadOnly(Of T)

        Public Event ValueChanged

        Private ReadOnly _callback As String

        Public Sub New(value As T, callback As String, name As AttributeName , category As AttributeCategory)
            MyBase.New(value, name, category)
            _callback = callback
        End Sub

        Sub SetValue(newValue As T)
            If Not Value.Equals(newValue) Then
                Value = newValue
                RaiseEvent ValueChanged
            End If
        End Sub

        Public Overrides ReadOnly Property Xml As String
            Get
                Return String.Format(XmlTemplate, Name.CamelCase(), _callback)
            End Get
        End Property
    End Class
End NameSpace