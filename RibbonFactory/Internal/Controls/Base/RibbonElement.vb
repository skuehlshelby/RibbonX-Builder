Imports System.ComponentModel
Imports RibbonX.Controls.Attributes
Imports RibbonX.InternalApi
Imports RibbonX.Utilities

Namespace Controls.Base

    Friend MustInherit Class RibbonElement
        Implements IRibbonElement
        Implements IAttributeSource

        Protected ReadOnly Attributes As IPropertyCollection
        Private ReadOnly changedProperty As PropertyChangedHelper = New PropertyChangedHelper()
        Private ReadOnly neededRefresh As RefreshNeededHelper = New RefreshNeededHelper()
        Private ReadOnly refreshSuspensionTracker As RefreshSuspensionTracker = New RefreshSuspensionTracker()

        Protected Sub New(attributes As IPropertyCollection, Optional tag As Object = Nothing)
            Me.Attributes = NotNull(attributes)
            Me.Tag = tag
            AddHandler attributes.PropertyChanged, AddressOf BubblePropertyChanged
            AddHandler refreshSuspensionTracker.HitZero, AddressOf EmptyMessageQueues
        End Sub

        Public Function SuspendRefreshing() As IDisposable Implements IRibbonElement.SuspendRefreshing
            Return refreshSuspensionTracker.AddSuspension()
        End Function

        Private Sub BubblePropertyChanged(sender As Object, e As PropertyChangedEventArgs)
            RaiseEvent PropertyChanged(Me, e)
            RaiseEvent RefreshNeeded(Me, New RefreshNeededEventArgs(ID))
        End Sub

        Protected Sub OnPropertyChanged(sender As Object, e As PropertyChangedEventArgs)
            RaiseEvent PropertyChanged(Me, e)
        End Sub

        Protected Sub OnRefreshNeeded(sender As Object, e As RefreshNeededEventArgs)
            RaiseEvent RefreshNeeded(Me, e)
        End Sub

        Protected Sub TriggerRefresh()
            RaiseEvent RefreshNeeded(Me, New RefreshNeededEventArgs(ID))
        End Sub

        Private Sub EmptyMessageQueues(sender As Object, e As EventArgs)
            changedProperty.EmptyQueue(Me)
            neededRefresh.EmptyQueue(Me)
        End Sub


        Public Custom Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
            AddHandler(value As PropertyChangedEventHandler)
                changedProperty.Handlers.Add(value)
            End AddHandler
            RemoveHandler(value As PropertyChangedEventHandler)
                changedProperty.Handlers.Remove(value)
            End RemoveHandler
            RaiseEvent(sender As Object, e As PropertyChangedEventArgs)
                If refreshSuspensionTracker.HasActiveRefreshSuspension Then
                    changedProperty.EventQueue.Enqueue(e)
                Else
                    changedProperty.InvokeAllHandlers(Me, e)
                End If
            End RaiseEvent
        End Event

        Public Custom Event RefreshNeeded As EventHandler(Of RefreshNeededEventArgs) Implements IRibbonElement.RefreshNeeded
            AddHandler(value As EventHandler(Of RefreshNeededEventArgs))
                neededRefresh.Handlers.Add(value)
            End AddHandler
            RemoveHandler(value As EventHandler(Of RefreshNeededEventArgs))
                neededRefresh.Handlers.Remove(value)
            End RemoveHandler
            RaiseEvent(sender As Object, e As RefreshNeededEventArgs)
                If refreshSuspensionTracker.HasActiveRefreshSuspension Then
                    neededRefresh.EventQueue.Enqueue(e)
                Else
                    neededRefresh.InvokeAllHandlers(Me, e)
                End If
            End RaiseEvent
        End Event

        Public ReadOnly Property ID As String Implements IRibbonElement.Id
            Get
                Return Attributes.Get(Category.IdType)
            End Get
        End Property

        Public Property Tag As Object Implements IRibbonElement.Tag

        Public MustOverride Function Clone() As Object Implements ICloneable.Clone

        Private Function GetAttributes() As IPropertyCollection Implements IAttributeSource.GetAttributes
            Return CType(Attributes.Clone(), IPropertyCollection)
        End Function

        Public Overridable Function ToXml(Optional exclude As ExcludedAttributes = ExcludedAttributes.None) As String Implements IRibbonElement.ToXml
            Return $"<{[GetType]().Name.CamelCase()} {Attributes.ToXml(exclude)}/>"
        End Function

#Region "Overrides and Equality Comparison"

        Public Overloads Overrides Function Equals(obj As Object) As Boolean
            Return Equals(TryCast(obj, IRibbonElement))
        End Function

        Public Overloads Function Equals(other As IRibbonElement) As Boolean Implements IEquatable(Of IRibbonElement).Equals
            Return other IsNot Nothing AndAlso other.Id.Equals(ID)
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return ID.GetHashCode()
        End Function

        Public Overrides Function ToString() As String
            Return $"{[GetType]().Name} (ID:{ID})"
        End Function

#End Region

    End Class

End Namespace