Imports RibbonFactory.Utilities

Namespace RibbonAttributes
	''' <summary>
	''' Provides a mechanism by which two attributes can share the same value.
	''' </summary>
	''' <typeparam name="T"></typeparam>
	Friend Class RibbonAttributeWrappedValue(Of T)
		Implements IRibbonAttributeReadWrite(Of T)
		Implements ICloneable

		Public Event ValueChanged() Implements IRibbonAttribute.ValueChanged

		Private Const XmlTemplate As String = "{0}=""{1}"""

		Private ReadOnly _value As ValueWrapper(Of T)
		Private ReadOnly _callback As String

		Public Sub New(value As ValueWrapper(Of T), callback As String, name As AttributeName, category As AttributeCategory)
			_value = value
			_callback = callback
			Me.Name = name
			Me.Category = category
		End Sub

		Public ReadOnly Property Name As AttributeName Implements IRibbonAttribute.Name

		Public ReadOnly Property Category As AttributeCategory Implements IRibbonAttribute.Category

		Public Function GetValue() As T Implements IRibbonAttributeReadOnly(Of T).GetValue
			Return _value.Value
		End Function

		Public Sub SetValue(value As T) Implements IRibbonAttributeReadWrite(Of T).SetValue
			If Not _value.Value.Equals(value) Then
				_value.Value = value
				RaiseEvent ValueChanged()
			End If
		End Sub

		Public Overrides Function ToString() As String
			Return String.Format(XmlTemplate, Name.CamelCase(), _callback)
		End Function

		Public Function Clone() As Object Implements ICloneable.Clone
			Return New RibbonAttributeWrappedValue(Of T)(_value.Clone(), String.Copy(_callback), Name.Clone(), Category.Clone())
		End Function
	End Class

End Namespace