Namespace RibbonAttributes.Categories.Pressed

    Public Class GetPressed
        Inherits RibbonAttribute(Of Boolean, GetPressed)

        Private ReadOnly _callback As String

        Public Sub New(value As Boolean, callback As FromControl(Of Boolean))
            MyBase.New(value)
            _callback = callback.Method.Name
        End Sub

        Public Sub SetValue(value As Boolean)
            MyBase.Value = value
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(GetPressed), _callback)
            End Get
        End Property
    End Class

End Namespace
