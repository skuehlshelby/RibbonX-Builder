Public Structure KeyTip
    Private ReadOnly _value As String

    Private Sub New (value As String)
        _value = value
    End Sub

    Public Overrides Function ToString() As String
        Return _value
    End Function

    Public Overrides Function Equals(obj As Object) As Boolean
        Return _value.Equals(obj)
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return _value.GetHashCode()
    End Function

    Public Shared Widening Operator CType(value As String) As KeyTip
        Return New KeyTip(Validate(value))
    End Operator

    Private Shared Function Validate(value As String) As String
        If value.Length < 1 OrElse value.Length > 3 Then
            Throw New ArgumentException($"KeyTips must be between one and three characters. Value {value} is invalid.")
        End If

        If value.ToCharArray().Any(Function(letter) Not Char.IsLetterOrDigit(letter)) Then
            Throw New ArgumentException($"KeyTips must be alphanumeric. Value {value} is invalid.")
        End If

        If Not KeyTipsInUse.Add(value.ToUpper()) Then
            Throw New ArgumentException($"KeyTips must be unique. Value {value} is already in use.")
        End If

        Return value.ToUpper()
    End Function

    Private Shared ReadOnly KeyTipsInUse As ISet(Of String) = New HashSet(Of String)
End Structure