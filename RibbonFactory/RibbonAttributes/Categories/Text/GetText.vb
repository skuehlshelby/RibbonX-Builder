
Namespace RibbonAttributes.Categories.Text

    Public Class GetText
        Inherits RibbonAttribute(Of String, GetText)

        Private ReadOnly _callback As String

        Public Sub New(value As String, callback As FromControl(Of String))
            MyBase.New(value)
            _callback = callback.Method.Name
        End Sub

        Public Sub SetValue(value As String)
            Me.Value = value
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(GetText)), _callback)
            End Get
        End Property
    End Class
End NameSpace