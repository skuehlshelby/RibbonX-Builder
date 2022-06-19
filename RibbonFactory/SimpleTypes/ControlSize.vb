Imports RibbonX.Utilities

Namespace SimpleTypes

    Public NotInheritable Class ControlSize
        Inherits Enumeration

        Private Sub New(value As Integer, name As String)
            MyBase.New(value, name)
        End Sub

        Public Shared ReadOnly Property Normal As ControlSize = New ControlSize(0, "normal")

        Public Shared ReadOnly Property Large As ControlSize = New ControlSize(1, "large")

        Public Overrides Function Clone() As Object
            Return New ControlSize(value, name)
        End Function

    End Class

End Namespace