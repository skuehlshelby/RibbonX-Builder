Imports RibbonFactory.Enums

Namespace RibbonAttributes.Categories.Size
    Friend Class GetSize
        Inherits Size

        Private ReadOnly Callback As String

        Public Sub New(Value As ControlSize, Callback As FromControl(Of ControlSize))
            MyBase.New(Value)
            Me.Callback = Callback.Method.Name
        End Sub

        Public Sub SetValue(Value As ControlSize)
            If Not Me.Value.Equals(Value) Then
                Me.Value = Value
            End If
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(GetSize), Callback)
            End Get
        End Property
    End Class
End Namespace