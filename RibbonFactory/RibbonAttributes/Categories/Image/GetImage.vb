Namespace RibbonAttributes.Categories.Image
    Friend Class GetImage
        Inherits ImageBase

        Private ReadOnly Callback As String
        Private Value As stdole.IPictureDisp

        Public Sub New(Value As stdole.IPictureDisp, Callback As FromControl(Of stdole.IPictureDisp))
            Me.Value = Value
            Me.Callback = Callback.Method.Name
        End Sub

        Public Function GetValue() As stdole.IPictureDisp
            Return Value
        End Function

        Public Sub SetValue(Value As stdole.IPictureDisp)
            If Not Me.Value.Equals(Value) Then
                Me.Value = Value
            End If
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(GetImage), Callback)
            End Get
        End Property
    End Class
End Namespace