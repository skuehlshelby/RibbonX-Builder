Namespace RibbonAttributes.Categories.Enabled
    Friend NotInheritable Class GetEnabled
        Inherits Enabled
        Implements IMutableRibbonAttribute(Of Boolean)

        Private ReadOnly _callback As String

        Public Sub New(value As Boolean, callback As FromControl(Of Boolean))
            MyBase.New(Value)
            Me._callback = Callback.Method.Name
        End Sub

        Public Sub SetValue(value As Boolean) Implements IMutableRibbonAttribute(Of Boolean).SetValue
            If Not Me.Value.Equals(Value) Then
                Me.Value = Value
            End If
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(GetEnabled), _callback)
            End Get
        End Property
    End Class
End Namespace