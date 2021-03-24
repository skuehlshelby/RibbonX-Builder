Namespace RibbonAttributes.Categories.OnAction
    Public NotInheritable Class OnAction
        Inherits RibbonAttribute(Of Action, OnAction)
        Implements IRibbonAttribute(Of Action)

        Private ReadOnly _callback As String

        Public Sub New(value As Action, callback As RibbonFactory.OnAction)
            MyBase.New(value)
            _callback = callback.Method.Name
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(OnAction)), _callback)
            End Get
        End Property
    End Class
End Namespace
