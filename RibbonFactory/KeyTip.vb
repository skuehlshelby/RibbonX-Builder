Public Structure KeyTip
	Implements IEquatable(Of KeyTip)
	Private ReadOnly _value As String

	Private Sub New(value As String)
		_value = value

		If AddToKeyTipsInUse Then
			KeyTipsInUse.Add(Me)
		End If
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

	Public Shared Widening Operator CType(value As Integer) As KeyTip
		Dim baseChars As Char() = "ABCDEFGHIJKLMNOPQRSTUVWXYZ123456789".ToCharArray()
		Dim radix As Integer = baseChars.Length
		Dim result As String = String.Empty

		Do
			value -= 1
			result = (baseChars(value Mod radix)) & result
			value \= radix
		Loop While value > 0

		Return New KeyTip(Validate(result))
	End Operator

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

	Private Shared ReadOnly KeyTipsInUse As ISet(Of KeyTip) = New HashSet(Of KeyTip)()

	Private Shared Seed As Integer = 1
	Private Shared AddToKeyTipsInUse As Boolean

	Public Shared Function NextAvailable() As KeyTip
		Try
			AddToKeyTipsInUse = False

			While KeyTipsInUse.Contains(Seed)
				Seed += 1
			End While

		Finally
			AddToKeyTipsInUse = True
		End Try

		Return Seed
	End Function

End Structure