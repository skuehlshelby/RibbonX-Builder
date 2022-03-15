Namespace RibbonAttributes

	<DebuggerDisplay("{GetDebuggerDisplay(),nq}")>
	Friend Class RibbonAttributeInvocationList
		Implements IRibbonAttributeReadOnly(Of ICollection(Of EventHandler))

		Public Event ValueChanged() Implements IRibbonAttribute.ValueChanged

		Private ReadOnly _value As ICollection(Of EventHandler)

		Public Sub New(handlers As ICollection(Of EventHandler), name As AttributeName, category As AttributeCategory)
			_value = handlers
			Me.Name = name
			Me.Category = category
		End Sub

		Public ReadOnly Property Name As AttributeName Implements IRibbonAttribute.Name

		Public ReadOnly Property Category As AttributeCategory Implements IRibbonAttribute.Category

		Public Function GetValue() As ICollection(Of EventHandler) Implements IRibbonAttributeReadOnly(Of ICollection(Of EventHandler)).GetValue
			Return _value
		End Function

		Public Overrides Function ToString() As String
			Return String.Empty
		End Function

		Private Function GetDebuggerDisplay() As String
			Return $"{NameOf(RibbonAttributeInvocationList)} ({Name}): {_value.Count} subscribers"
		End Function

	End Class


	<DebuggerDisplay("{GetDebuggerDisplay(),nq}")>
	Friend Class RibbonAttributeInvocationList(Of T)
		Implements IRibbonAttributeReadOnly(Of ICollection(Of EventHandler(Of T)))

		Public Event ValueChanged() Implements IRibbonAttribute.ValueChanged

		Private ReadOnly _value As ICollection(Of EventHandler(Of T))

		Public Sub New(handlers As ICollection(Of EventHandler(Of T)), name As AttributeName, category As AttributeCategory)
			_value = handlers
			Me.Name = name
			Me.Category = category
		End Sub

		Public ReadOnly Property Name As AttributeName Implements IRibbonAttribute.Name

		Public ReadOnly Property Category As AttributeCategory Implements IRibbonAttribute.Category

		Public Function GetValue() As ICollection(Of EventHandler(Of T)) Implements IRibbonAttributeReadOnly(Of ICollection(Of EventHandler(Of T))).GetValue
			Return _value
		End Function

		Public Overrides Function ToString() As String
			Return String.Empty
		End Function

		Private Function GetDebuggerDisplay() As String
			Return $"{NameOf(RibbonAttributeInvocationList(Of T))} ({Name}): {_value.Count} subscribers"
		End Function

	End Class

End Namespace

