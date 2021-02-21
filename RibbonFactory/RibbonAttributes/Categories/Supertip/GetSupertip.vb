Namespace RibbonAttributes.Categories.Supertip
    Friend NotInheritable Class GetSupertip
        Inherits Supertip

        Private ReadOnly Callback As String

        Public Sub New(Value As String, Callback As FromControl(Of String))
            MyBase.New(Value)
            Me.Callback = Callback.Method.Name
        End Sub

        Public Sub SetValue(Value As String)
            If Not Me.Value.Equals(Value) Then
                Me.Value = Value
            End If
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(GetSupertip), Callback)
            End Get
        End Property
    End Class
End Namespace