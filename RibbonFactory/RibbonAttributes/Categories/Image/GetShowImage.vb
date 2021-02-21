Namespace RibbonAttributes.Categories.Image
    Friend Class GetShowImage
        Inherits ShowImage

        Private ReadOnly _callback As String

        Public Sub New(value As Boolean, callback As FromControl(Of Boolean))
            MyBase.New(value)
            _callback = callback.Method.Name
        End Sub

        Public Sub SetValue(value As Boolean)
            If Not Me.Value.Equals(value) Then
                Me.Value = value
            End If
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(GetShowImage), _callback)
            End Get
        End Property
    End Class
End NameSpace