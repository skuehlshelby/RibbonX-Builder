Namespace Enums.ImageMSO
    Public NotInheritable Class Number
        Inherits ImageMSO
        
        Private Sub New(name As String)
            MyBase.New(name)
        End Sub
        
        Public Shared ReadOnly Property Zero As Number = New Number("_0")

        Public Shared ReadOnly Property One As Number = New Number("_1")

        Public Shared ReadOnly Property Two As Number = New Number("_2")

        Public Shared ReadOnly Property Three As Number = New Number("_3")

        Public Shared ReadOnly Property Four As Number = New Number("_4")

        Public Shared ReadOnly Property Five As Number = New Number("_5")

        Public Shared ReadOnly Property Six As Number = New Number("_6")

        Public Shared ReadOnly Property Seven As Number = New Number("_7")

        Public Shared ReadOnly Property Eight As Number = New Number("_8")

        Public Shared ReadOnly Property Nine As Number = New Number("_9")

    End Class
End NameSpace