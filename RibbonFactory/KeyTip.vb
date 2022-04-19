Public Structure KeyTip
	Implements IEquatable(Of KeyTip)
	Private ReadOnly _value As String

	Private Sub New(value As String)
		_value = value
	End Sub

	Public Overrides Function ToString() As String
		Return _value
	End Function

	Public Overrides Function Equals(obj As Object) As Boolean
		Return _value.Equals(obj.ToString(), StringComparison.OrdinalIgnoreCase)
	End Function

	Public Overloads Function Equals(other As KeyTip) As Boolean Implements IEquatable(Of KeyTip).Equals
		Return other._value.Equals(_value, StringComparison.OrdinalIgnoreCase)
	End Function

	Public Overrides Function GetHashCode() As Integer
		Return _value.GetHashCode()
	End Function

	Public Shared Widening Operator CType(value As String) As KeyTip
		Return New KeyTip(Validate(value))
	End Operator

	Public Shared Widening Operator CType(value As KeyTip) As String
		Return value.ToString()
	End Operator

	Private Shared Function Validate(value As String) As String
		If value.Length < 1 OrElse value.Length > 3 Then
			Throw New FormatException($"KeyTips must be between one and three characters. Value {value} is invalid.")
		End If

		If value.ToCharArray().Any(Function(letter) Not Char.IsLetterOrDigit(letter)) Then
			Throw New FormatException($"KeyTips must be alphanumeric. Value {value} is invalid.")
		End If

		If value.ToCharArray(0, value.Length - 1).Any(Function(letter) Char.IsDigit(letter)) Then
			Throw New FormatException($"Only the last character of a keytip may be numeric. Value {value} is invalid.")
		End If

		Return value
	End Function

End Structure