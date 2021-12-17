Public MustInherit Class RibbonElement
    Implements IEquatable(Of RibbonElement)
    
    Public Event ValueChanged(sender As Object, e As ValueChangedEventArgs)
    Private _refreshAutomatically As Boolean

    Protected Sub RefreshNeeded()
        If _refreshAutomatically Then
            RaiseEvent ValueChanged(Me, New ValueChangedEventArgs(ID))
        End If
    End Sub

    Protected Friend Sub New(Optional tag As Object = Nothing)
        Me.Tag = tag
        _refreshAutomatically = True
    End Sub

    Public MustOverride ReadOnly Property ID As String

    Public Property Tag As Object

    Public MustOverride ReadOnly Property XML As String

    Public Sub SuspendAutomaticRefreshing()
        _refreshAutomatically = False
    End Sub

    Public Sub ResumeAutomaticRefreshing(Optional triggerRefreshNow As Boolean = True)
        _refreshAutomatically = True

        If triggerRefreshNow Then
            RefreshNeeded()
        End If
    End Sub

#Region "Overrides and Equality Comparison"

    Public Overrides Function ToString() As String
        Return XML
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return ID.GetHashCode()
    End Function

    Public Overloads Overrides Function Equals(obj As Object) As Boolean
        Return Equals(TryCast(obj, RibbonElement))
    End Function
    
    Public Overloads Function Equals(other As RibbonElement) As Boolean Implements IEquatable(Of RibbonElement).Equals
        Return other IsNot Nothing AndAlso other.ID.Equals(ID)
    End Function

#End Region

End Class
