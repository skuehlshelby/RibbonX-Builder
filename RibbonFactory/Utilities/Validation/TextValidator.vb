Namespace Utilities.Validation
    
    Friend Class TextValidator
        Implements IValidate(Of String)
        Implements IValidationResult

        Private ReadOnly _rule As Predicate(Of String)
        Private _success As Boolean

        Public Sub New (rule As Predicate(Of String), failureMessage As String)
            Require(Of ArgumentNullException)(rule IsNot Nothing)
            _rule = rule
            Me.FailureMessage = failureMessage
        End Sub

        Public Function Validate(value As String) As IValidationResult Implements IValidate(Of String).Validate
            _success = _rule.Invoke(value)
            Return Me
        End Function

        Private ReadOnly Property Success As Boolean Implements IValidationResult.Success
            Get
                Return _success
            End Get
        End Property

        Private ReadOnly Property FailureMessage As String Implements IValidationResult.FailureMessage

    End Class

End NameSpace