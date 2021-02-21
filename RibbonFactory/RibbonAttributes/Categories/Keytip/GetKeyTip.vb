Namespace RibbonAttributes.Categories.KeyTip
    Friend NotInheritable Class GetKeyTip
        Inherits KeyTip

        Private ReadOnly _callback As String

        Public Sub New(value As Char, callback As FromControl(Of Char))
            MyBase.New(value)
            _callback = callback.Method.Name
        End Sub

        Public Sub SetValue(value As Char)
            If Not Me.Value.Equals(value) Then
                Me.Value = value
            End If
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(GetKeyTip), _callback)
            End Get
        End Property
    End Class
End Namespace