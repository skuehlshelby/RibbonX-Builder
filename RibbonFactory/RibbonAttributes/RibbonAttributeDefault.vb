Imports RibbonX.Utilities

Namespace RibbonAttributes

	<DebuggerDisplay("{GetDebuggerDisplay(),nq}")>
	Friend Class RibbonAttributeDefault(Of T)
		Implements IRibbonAttributeReadOnly(Of T)
		Implements ICloneable

		Public Event ValueChanged() Implements IRibbonAttribute.ValueChanged

		Private Const XmlTemplate As String = "{0}=""{1}"""

		Private ReadOnly value As T

		Public Sub New(value As T, name As AttributeName, category As AttributeCategory)
			Me.value = value
			Me.Name = name
			Me.Category = category
		End Sub

		Public ReadOnly Property Name As AttributeName Implements IRibbonAttribute.Name

		Public ReadOnly Property Category As AttributeCategory Implements IRibbonAttribute.Category

		Public Function GetValue() As T Implements IRibbonAttributeReadOnly(Of T).GetValue
			Return value
		End Function

		Public Overrides Function ToString() As String
			Return String.Empty
		End Function

		Private Function GetDebuggerDisplay() As String
			If TypeOf value Is Boolean Then
				Return String.Format(XmlTemplate, Name.CamelCase(), value.CamelCase())
			Else
				Return String.Format(XmlTemplate, Name.CamelCase(), value)
			End If
		End Function

		Private Function Clone() As Object Implements ICloneable.Clone
			Dim valueCopy As T = If(TypeOf value Is ICloneable, DirectCast(DirectCast(value, ICloneable).Clone(), T), value)

			Return New RibbonAttributeDefault(Of T)(valueCopy, Name.Clone(), Category.Clone())
		End Function
	End Class

End Namespace