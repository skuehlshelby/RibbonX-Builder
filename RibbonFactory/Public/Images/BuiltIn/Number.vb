Namespace Images.BuiltIn
    Public NotInheritable Class Number
        Inherits ImageMSO
        Private Sub New(value As Integer, name As String)
            MyBase.New(value, name)
        End Sub

        Public ReadOnly Property NumericValue As Integer
            Get
                Return value
            End Get
        End Property

        Public Shared ReadOnly Property Zero As Number = New Number(0, "_0")
        Public Shared ReadOnly Property One As Number = New Number(1, "_1")
        Public Shared ReadOnly Property Two As Number = New Number(2, "_2")
        Public Shared ReadOnly Property Three As Number = New Number(3, "_3")
        Public Shared ReadOnly Property Four As Number = New Number(4, "_4")
        Public Shared ReadOnly Property Five As Number = New Number(5, "_5")
        Public Shared ReadOnly Property Six As Number = New Number(6, "_6")
        Public Shared ReadOnly Property Seven As Number = New Number(7, "_7")
        Public Shared ReadOnly Property Eight As Number = New Number(8, "_8")
        Public Shared ReadOnly Property Nine As Number = New Number(9, "_9")

        Public Overrides Function Clone() As Object
            Return New Number(value, name)
        End Function

    End Class

End Namespace