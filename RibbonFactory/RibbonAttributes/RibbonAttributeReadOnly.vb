Imports RibbonFactory.Utilities

Namespace RibbonAttributes

	Friend Class RibbonAttributeReadOnly(Of T)
		Implements IRibbonAttributeReadOnly(Of T)

		Public Event ValueChanged() Implements IRibbonAttribute.ValueChanged

		Private Const XmlTemplate As String = "{0}=""{1}"""

		Private ReadOnly _value As T

		Public Sub New(value As T, name As AttributeName, category As AttributeCategory)
			Require(Of ArgumentNullException)(name IsNot Nothing)
			Require(Of ArgumentNullException)(category IsNot Nothing)

			_value = value
			Me.Name = name
			Me.Category = category
		End Sub

		Public ReadOnly Property Name As AttributeName Implements IRibbonAttribute.Name

		Public ReadOnly Property Category As AttributeCategory Implements IRibbonAttribute.Category

		Public Function GetValue() As T Implements IRibbonAttributeReadOnly(Of T).GetValue
			Return _value
		End Function

		Public Overrides Function ToString() As String
			If TypeOf _value Is Boolean Then
				Return String.Format(XmlTemplate, Name.CamelCase(), _value.CamelCase())
			Else
				Return String.Format(XmlTemplate, Name.CamelCase(), _value)
			End If
		End Function

	End Class

End Namespace