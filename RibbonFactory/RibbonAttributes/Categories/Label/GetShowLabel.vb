Namespace RibbonAttributes.Categories.Label
    Friend NotInheritable Class GetShowLabel
        Inherits ShowLabel

        Private ReadOnly Callback As String

        Public Sub New(Value As Boolean, Callback As FromControl(Of Boolean))
            MyBase.New(Value)
            Me.Callback = Callback.Method.Name
        End Sub

        Public Sub SetValue(Value As Boolean)
            If Not Me.Value.Equals(Value) Then
                Me.Value = Value
            End If
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(GetShowLabel)), Callback)
            End Get
        End Property
    End Class
End Namespace
