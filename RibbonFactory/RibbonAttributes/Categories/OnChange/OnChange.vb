Namespace RibbonAttributes.Categories.OnChange
    Public Class OnChange
        Inherits RibbonAttribute(Of Action, OnChange)

        Private ReadOnly _callback As String

        Public Sub New(value As Action, callback As RibbonFactory.OnAction)
            MyBase.New(value)
            _callback = callback.Method.Name
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(OnChange)), _callback)
            End Get
        End Property
    End Class
End NameSpace