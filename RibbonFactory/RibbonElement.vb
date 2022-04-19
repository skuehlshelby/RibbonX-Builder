Public MustInherit Class RibbonElement
	Implements IEquatable(Of RibbonElement)

	Private _refreshEnabled As Boolean = True

	Protected Sub New(Optional tag As Object = Nothing)
		Me.Tag = tag
	End Sub

	Public Event ValueChanged As EventHandler(Of ValueChangedEventArgs)

	Public MustOverride ReadOnly Property ID As String

	Public Property Tag As Object

	Public MustOverride ReadOnly Property XML As String

#Region "Refreshing"

	Protected Sub RefreshNeeded()
		If _refreshEnabled Then
			RaiseEvent ValueChanged(Me, New ValueChangedEventArgs(ID))
		End If
	End Sub

	Public Function RefreshSuspension(Optional refreshOnDispose As Boolean = True) As IDisposable
		Return New UpdateBlock(Me, refreshOnDispose)
	End Function

	Private NotInheritable Class UpdateBlock
		Implements IDisposable

		Private parent As RibbonElement
		Private refreshOnDispose As Boolean

		Public Sub New(parent As RibbonElement, refreshOnDispose As Boolean)
			Me.parent = parent
			Me.refreshOnDispose = refreshOnDispose
			Me.parent._refreshEnabled = False
		End Sub

		Private Sub Dispose() Implements IDisposable.Dispose
			parent._refreshEnabled = True

			If refreshOnDispose Then
				parent.RefreshNeeded()
			End If

			parent = Nothing
		End Sub
	End Class

#End Region

#Region "Overrides and Equality Comparison"

	Public Overloads Overrides Function Equals(obj As Object) As Boolean
		Return Equals(TryCast(obj, RibbonElement))
	End Function

	Public Overloads Function Equals(other As RibbonElement) As Boolean Implements IEquatable(Of RibbonElement).Equals
		Return other IsNot Nothing AndAlso other.ID.Equals(ID)
	End Function

	Public Overrides Function GetHashCode() As Integer
		Return ID.GetHashCode()
	End Function

	Public Overrides Function ToString() As String
		Return XML
	End Function

#End Region

End Class