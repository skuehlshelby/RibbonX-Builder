Namespace RibbonAttributes

	''' <summary>
	''' This class encapsulates the behavior of a property which can be set either programmatically
	''' or from the the Office UI and therefore uses two callbacks to modify the same value.
	''' </summary>
	''' <typeparam name="T"></typeparam>
	Friend Class RibbonAttributeComposite(Of T)
		Implements IRibbonAttributeReadWrite(Of T)

		Public Event ValueChanged() Implements IRibbonAttribute.ValueChanged

		Private Const XmlTemplate As String = "{0}=""{1}"""

		Private ReadOnly _nameOne As AttributeName
		Private ReadOnly _nameTwo As AttributeName
		Private ReadOnly _categoryOne As AttributeCategory
		Private ReadOnly _categoryTwo As AttributeCategory
		Private ReadOnly _callbackOne As String
		Private ReadOnly _callbackTwo As String
		Private _value As T

		Public Sub New(value As T, nameOne As AttributeName, categoryOne As AttributeCategory, callbackOne As String, nameTwo As AttributeName, categoryTwo As AttributeCategory, callbackTwo As String)
			Utilities.Require(Of ArgumentNullException)(value IsNot Nothing, $"Argument '{NameOf(value)}' cannot be null.")
			Utilities.Require(Of ArgumentNullException)(nameOne IsNot Nothing, $"Argument '{NameOf(nameOne)}' cannot be null.")
			Utilities.Require(Of ArgumentNullException)(categoryOne IsNot Nothing, $"Argument '{NameOf(categoryOne)}' cannot be null.")
			Utilities.Require(Of ArgumentNullException)(Not String.IsNullOrWhiteSpace(callbackOne), $"Argument '{NameOf(callbackOne)}' cannot be null or whitespace.")
			Utilities.Require(Of ArgumentNullException)(nameTwo IsNot Nothing, $"Argument '{NameOf(nameTwo)}' cannot be null.")
			Utilities.Require(Of ArgumentNullException)(categoryTwo IsNot Nothing, $"Argument '{NameOf(categoryTwo)}' cannot be null.")
			Utilities.Require(Of ArgumentNullException)(Not String.IsNullOrWhiteSpace(callbackTwo), $"Argument '{NameOf(callbackTwo)}' cannot be null or whitespace.")

			_value = value
			_nameOne = nameOne
			_categoryOne = categoryOne
			_callbackOne = callbackOne
			_nameTwo = nameTwo
			_categoryTwo = categoryTwo
			_callbackTwo = callbackTwo
		End Sub

		Public Function GetValue() As T Implements IRibbonAttributeReadOnly(Of T).GetValue
			Return _value
		End Function

		Public Sub SetValue(value As T) Implements IRibbonAttributeReadWrite(Of T).SetValue
			If Not _value.Equals(value) Then
				_value = value
				RaiseEvent ValueChanged()
			End If
		End Sub

		Public Function IsNamed(name As AttributeName) As Boolean Implements IRibbonAttribute.IsNamed
			Return name IsNot Nothing AndAlso (name.Equals(_nameOne) OrElse name.Equals(_nameTwo))
		End Function

		Public Function IsExclusiveWith(name As AttributeName) As Boolean Implements IRibbonAttribute.IsExclusiveWith
			Return name IsNot Nothing AndAlso (_categoryOne.Contains(name) OrElse _categoryTwo.Contains(name))
		End Function

		Public Function IsExclusiveWith(category As AttributeCategory) As Boolean Implements IRibbonAttribute.IsExclusiveWith
			Return category IsNot Nothing AndAlso (_categoryOne.Equals(category) OrElse _categoryTwo.Equals(category))
		End Function

#Region "Overrides and Equality Comparison"

		Public Overrides Function ToString() As String
			Return String.Join(" ", String.Format(XmlTemplate, _nameOne, _callbackOne), String.Format(XmlTemplate, _nameTwo, _callbackTwo))
		End Function

		Public Overrides Function Equals(obj As Object) As Boolean
			Return Equals(TryCast(obj, IRibbonAttribute))
		End Function

		Public Overloads Function Equals(other As IRibbonAttribute) As Boolean Implements IEquatable(Of IRibbonAttribute).Equals
			Return other IsNot Nothing AndAlso (other.IsExclusiveWith(_categoryOne) OrElse other.IsExclusiveWith(_categoryTwo))
		End Function

		Public Overrides Function GetHashCode() As Integer
			Return _categoryOne.GetHashCode() Xor _categoryTwo.GetHashCode()
		End Function

#End Region

	End Class

End Namespace