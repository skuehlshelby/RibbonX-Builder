Namespace RibbonAttributes.Categories.MaxLength
    
    Public Class MaxLength
        Inherits RibbonAttribute(Of UShort, MaxLength)

        Public Sub New(value As UShort)
            MyBase.New(value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(MaxLength)), Value)
            End Get
        End Property
    End Class

End NameSpace