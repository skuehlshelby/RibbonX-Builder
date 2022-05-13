Namespace Enums.ImageMSO
    Public NotInheritable Class Number
        Inherits ImageMSO
        Private Sub New(value As Integer, name As String)
            MyBase.New(value, name)
        End Sub

        Public Shared ReadOnly Property _0 As Number = New Number(1, NameOf(_0))
        Public Shared ReadOnly Property _1 As Number = New Number(2, NameOf(_1))
        Public Shared ReadOnly Property _2 As Number = New Number(3, NameOf(_2))
        Public Shared ReadOnly Property _3 As Number = New Number(4, NameOf(_3))
        Public Shared ReadOnly Property _4 As Number = New Number(5, NameOf(_4))
        Public Shared ReadOnly Property _5 As Number = New Number(6, NameOf(_5))
        Public Shared ReadOnly Property _6 As Number = New Number(7, NameOf(_6))
        Public Shared ReadOnly Property _7 As Number = New Number(8, NameOf(_7))
        Public Shared ReadOnly Property _8 As Number = New Number(9, NameOf(_8))
        Public Shared ReadOnly Property _9 As Number = New Number(10, NameOf(_9))

        Public Overrides Function Clone() As Object
            Return New Number(value, name)
        End Function

    End Class

End NameSpace