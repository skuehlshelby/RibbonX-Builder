Imports RibbonX.Utilities

Namespace SimpleTypes

    Public NotInheritable Class BoxStyle
        Inherits Enumeration

        Private Sub New(value As Integer, name As String)
            MyBase.New(value, name)
        End Sub

        Public Shared ReadOnly Property Horizontal As BoxStyle = New BoxStyle(0, "horizontal")

        Public Shared ReadOnly Property Vertical As BoxStyle = New BoxStyle(1, "vertical")

        Public Overrides Function Clone() As Object
            Return New BoxStyle(value, name)
        End Function

    End Class

End Namespace