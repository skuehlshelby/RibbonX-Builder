Namespace BuilderInterfaces

    Public Interface IValidateText(Of T)
    
        Function WithTextValidationRule(rule As Predicate(Of String)) As T

        Function WithTextValidationRule(rule As Predicate(Of String), failureMessage As String) As T

    End Interface

End NameSpace