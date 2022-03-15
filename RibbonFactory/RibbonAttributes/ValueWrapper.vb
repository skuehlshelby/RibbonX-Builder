Namespace RibbonAttributes

	Friend Class ValueWrapper(Of T)
		Implements ICloneable

		Public Sub New(value As T)
			Me.Value = value
		End Sub

		Public Property Value As T

		Public Function Clone() As Object Implements ICloneable.Clone
			Return New ValueWrapper(Of T)(Value)
		End Function
	End Class

End Namespace

