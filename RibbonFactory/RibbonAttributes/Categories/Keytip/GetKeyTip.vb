Namespace RibbonAttributes.Categories.KeyTip
    Friend NotInheritable Class GetKeyTip
        Inherits KeyTip

        Private ReadOnly _callback As String

        Public Sub New(value As RibbonFactory.KeyTip, callback As FromControl(Of RibbonFactory.KeyTip))
            MyBase.New(value)
            _callback = callback.Method.Name
        End Sub

        Public Sub SetValue(value As RibbonFactory.KeyTip)
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