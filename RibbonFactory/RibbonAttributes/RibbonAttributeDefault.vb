Imports RibbonFactory.Utilities

Namespace RibbonAttributes
	<DebuggerDisplay("{GetDebuggerDisplay(),nq}")>
	Friend Class RibbonAttributeDefault(Of T)
		Inherits RibbonAttributeReadOnly(Of T)

		Public Sub New(value As T, name As AttributeName, category As AttributeCategory)
			MyBase.New(value, name, category)
		End Sub

		Public Overrides ReadOnly Property Xml As String
			Get
				Return String.Empty
			End Get
		End Property

		Private Function GetDebuggerDisplay() As String
			If TypeOf Value Is Boolean Then
				Return String.Format(XmlTemplate, Name.CamelCase(), Value.CamelCase())
			Else
				Return String.Format(XmlTemplate, Name.CamelCase(), Value)
			End If
		End Function
	End Class

End Namespace