Namespace Utilities.Validation
    
    Public Interface IValidate(Of T)
    
        Function Validate(value As T) As IValidationResult

    End Interface

End NameSpace