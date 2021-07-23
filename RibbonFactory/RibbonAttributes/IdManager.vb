Namespace RibbonAttributes

    Friend NotInheritable Class IdManager
        Private Shared ReadOnly ControlIDs As IDictionary(Of Type, Integer) = New Dictionary(Of Type, Integer)

        Private Sub New()
        End Sub

        Public Shared Function GetID(Of T As RibbonElement)() As String
            Dim ribbonElementType As Type = GetType(T)

            If Not ControlIDs.ContainsKey(ribbonElementType) Then
                ControlIDs.Add(ribbonElementType, 0)
            End If

            ControlIDs.Item(ribbonElementType) += 1

            Return ribbonElementType.Name & ControlIDs.Item(ribbonElementType).ToString
        End Function
    End Class

End NameSpace