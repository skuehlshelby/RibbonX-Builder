Imports RibbonFactory.Utilities

Namespace RibbonAttributes
	Friend Class RibbonAttributeReadWrite(Of T)
		Implements IRibbonAttributeReadWrite(Of T)

		Public Event ValueChanged() Implements IRibbonAttribute.ValueChanged

		Private Const XmlTemplate As String = "{0}=""{1}"""

		Private ReadOnly _name As AttributeName
		Private ReadOnly _category As AttributeCategory
		Private ReadOnly _callback As String
		Private _value As T

		Public Sub New(value As T, name As AttributeName, category As AttributeCategory, callback As String)
			_value = value
			_name = name
			_category = category
			_callback = callback
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
			Return name IsNot Nothing AndAlso name.Equals(_name)
		End Function

		Public Function IsExclusiveWith(name As AttributeName) As Boolean Implements IRibbonAttribute.IsExclusiveWith
			Return name IsNot Nothing AndAlso _category.Contains(name)
		End Function

		Public Function IsExclusiveWith(category As AttributeCategory) As Boolean Implements IRibbonAttribute.IsExclusiveWith
			Return category IsNot Nothing AndAlso category.Equals(_category)
		End Function

#Region "Overrides and Equality Comparison"

		Public Overrides Function ToString() As String
			Return String.Format(XmlTemplate, _name.CamelCase(), _callback)
		End Function

		Public Overrides Function Equals(obj As Object) As Boolean
			Return Equals(TryCast(obj, IRibbonAttribute))
		End Function

		Public Overloads Function Equals(other As IRibbonAttribute) As Boolean Implements IEquatable(Of IRibbonAttribute).Equals
			Return other IsNot Nothing AndAlso other.IsExclusiveWith(_category)
		End Function

		Public Overrides Function GetHashCode() As Integer
			Return _category.GetHashCode()
		End Function

#End Region

	End Class

End Namespace