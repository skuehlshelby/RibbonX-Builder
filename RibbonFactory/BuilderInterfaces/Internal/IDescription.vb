Namespace BuilderInterfaces

    Public Interface IDescription(Of T)

        Function WithDescription(description As String) As T

        Function WithDescription(description As String, callback As FromControl(Of String)) As T

    End Interface

End NameSpace