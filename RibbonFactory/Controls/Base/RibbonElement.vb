Imports System.ComponentModel
Imports RibbonX.Api
Imports RibbonX.Controls.Attributes
Imports RibbonX.Utilities

Namespace Controls.Base

    Friend MustInherit Class RibbonElement
        Implements IRibbonElement

        Protected ReadOnly Attributes As AttributeSet
        Private ReadOnly bindings As ISet(Of IBinding)
        Private refreshState As IRefreshBehavior

        Protected Sub New(attributes As AttributeSet, Optional tag As Object = Nothing)
            Me.Attributes = attributes
            Me.Tag = tag
            refreshState = New RefreshEnabled(Me)
        End Sub

        Public Custom Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
            AddHandler(value As PropertyChangedEventHandler)
                refreshState.AddHandler(value)
            End AddHandler
            RemoveHandler(value As PropertyChangedEventHandler)
                refreshState.RemoveHandler(value)
            End RemoveHandler
            RaiseEvent(sender As Object, e As PropertyChangedEventArgs)
                refreshState.RaiseEvent(sender, e)
            End RaiseEvent
        End Event

        Public Custom Event ValueChanged As EventHandler(Of ValueChangedEventArgs)
            AddHandler(value As EventHandler(Of ValueChangedEventArgs))
                refreshState.AddHandler(value)
            End AddHandler
            RemoveHandler(value As EventHandler(Of ValueChangedEventArgs))
                refreshState.RemoveHandler(value)
            End RemoveHandler
            RaiseEvent(sender As Object, e As ValueChangedEventArgs)
                refreshState.RaiseEvent(sender, e)
            End RaiseEvent
        End Event

        Public ReadOnly Property ID As String Implements IRibbonElement.Id
            Get
                Return Attributes.Get(Of String)(AttributeCategory.IdType)
            End Get
        End Property

        Public Property Tag As Object Implements IRibbonElement.Tag

        Public Overridable Function ConvertToXml() As String Implements IRibbonElement.ConvertToXml
            Return $"<{[GetType]().Name.CamelCase()} { Attributes }/>"
        End Function

        Public Function GetChildren() As ICollection(Of IRibbonElement) Implements IRibbonElement.GetChildren
            Dim results As ICollection(Of IRibbonElement) = New LinkedList(Of IRibbonElement)()
            GetChildren(results)
            Return results.ToArray()
        End Function

        Protected Overridable Sub GetChildren(results As ICollection(Of IRibbonElement))

        End Sub

        Public MustOverride Function Clone() As Object Implements ICloneable.Clone

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
            Return ConvertToXml()
        End Function

#End Region

#Region "Binding"

        Public Function Bind(binding As IBinding) As IDisposable Implements IRibbonElement.Bind
            bindings.Add(binding)
            Return New BindingRemover
        End Function

        Private Class BindingRemover
            Implements IDisposable
            Private ReadOnly item As IBinding
            Private ReadOnly list As ICollection(Of IBinding)

            Public Sub Dispose() Implements IDisposable.Dispose
                list.Remove(item)
            End Sub
        End Class

#End Region

#Region "Refresh Suspension"

        Public Function RefreshSuspension() As IDisposable Implements IRibbonElement.SuspendRefreshing
            Return New Suspension(refreshState)
        End Function

        Private NotInheritable Class Suspension
            Implements IDisposable
            Private ReadOnly refreshState As IRefreshBehavior

            Public Sub New(refreshState As IRefreshBehavior)
                If refreshState Is Nothing Then
                    Throw New ArgumentNullException(NameOf(refreshState))
                End If

                Me.refreshState = refreshState
                refreshState.SuspendRefreshing()
            End Sub

            Public Sub Dispose() Implements IDisposable.Dispose
                refreshState.EnableRefreshing()
            End Sub
        End Class

        Private Interface IRefreshBehavior
            Sub [AddHandler](value As PropertyChangedEventHandler)
            Sub [AddHandler](value As EventHandler(Of ValueChangedEventArgs))
            Sub [RemoveHandler](value As PropertyChangedEventHandler)
            Sub [RemoveHandler](value As EventHandler(Of ValueChangedEventArgs))
            Sub [RaiseEvent](sender As Object, e As PropertyChangedEventArgs)
            Sub [RaiseEvent](sender As Object, e As ValueChangedEventArgs)
            Sub SuspendRefreshing()
            Sub EnableRefreshing()
        End Interface

        Private NotInheritable Class RefreshDisabled
            Implements IRefreshBehavior
            Private ReadOnly parent As RibbonElement
            Private ReadOnly events As EventHandlerList
            Private ReadOnly propertyChanges As Queue(Of Tuple(Of Object, PropertyChangedEventArgs)) = New Queue(Of Tuple(Of Object, PropertyChangedEventArgs))
            Private ReadOnly refreshes As Queue(Of Tuple(Of Object, ValueChangedEventArgs)) = New Queue(Of Tuple(Of Object, ValueChangedEventArgs))

            Public Sub New(parent As RibbonElement, events As EventHandlerList)
                If parent Is Nothing Then
                    Throw New ArgumentNullException(NameOf(parent))
                End If

                If events Is Nothing Then
                    Throw New ArgumentNullException(NameOf(events))
                End If

                Me.parent = parent
                Me.events = events
            End Sub

            Public Sub [RaiseEvent](sender As Object, e As PropertyChangedEventArgs) Implements IRefreshBehavior.RaiseEvent
                propertyChanges.Enqueue((sender, e).ToTuple())
            End Sub

            Public Sub [RaiseEvent](sender As Object, e As ValueChangedEventArgs) Implements IRefreshBehavior.RaiseEvent
                refreshes.Enqueue((sender, e).ToTuple())
            End Sub

            Public Sub [AddHandler](value As PropertyChangedEventHandler) Implements IRefreshBehavior.AddHandler
                events.AddHandler(NameOf(PropertyChanged), value)
            End Sub

            Public Sub [AddHandler](value As EventHandler(Of ValueChangedEventArgs)) Implements IRefreshBehavior.AddHandler
                events.AddHandler(NameOf(ValueChanged), value)
            End Sub

            Public Sub [RemoveHandler](value As PropertyChangedEventHandler) Implements IRefreshBehavior.RemoveHandler
                events.RemoveHandler(NameOf(PropertyChanged), value)
            End Sub

            Public Sub [RemoveHandler](value As EventHandler(Of ValueChangedEventArgs)) Implements IRefreshBehavior.RemoveHandler
                events.RemoveHandler(NameOf(ValueChanged), value)
            End Sub

            Public Sub SuspendRefreshing() Implements IRefreshBehavior.SuspendRefreshing

            End Sub

            Public Sub EnableRefreshing() Implements IRefreshBehavior.EnableRefreshing
                parent.refreshState = New RefreshEnabled(parent, events)

                While 0 < propertyChanges.Count
                    Dim sender As Object = Nothing
                    Dim e As PropertyChangedEventArgs = Nothing
                    propertyChanges.Dequeue().Deconstruct(sender, e)

                    parent.refreshState.RaiseEvent(sender, e)
                End While

                If 0 < refreshes.Count Then
                    Dim sender As Object = Nothing
                    Dim e As ValueChangedEventArgs = Nothing
                    refreshes.Dequeue().Deconstruct(sender, e)

                    parent.refreshState.RaiseEvent(sender, e)
                    refreshes.Clear()
                End If
            End Sub
        End Class

        Private NotInheritable Class RefreshEnabled
            Implements IRefreshBehavior
            Private ReadOnly parent As RibbonElement
            Private ReadOnly events As EventHandlerList

            Public Sub New(parent As RibbonElement)
                If parent Is Nothing Then
                    Throw New ArgumentNullException(NameOf(parent))
                End If

                Me.parent = parent
                events = New EventHandlerList()
            End Sub

            Public Sub New(parent As RibbonElement, events As EventHandlerList)
                If parent Is Nothing Then
                    Throw New ArgumentNullException(NameOf(parent))
                End If

                If events Is Nothing Then
                    Throw New ArgumentNullException(NameOf(events))
                End If

                Me.parent = parent
                Me.events = events
            End Sub

            Public Sub [RaiseEvent](sender As Object, e As PropertyChangedEventArgs) Implements IRefreshBehavior.RaiseEvent
                CType(events(NameOf(PropertyChanged)), PropertyChangedEventHandler).Invoke(sender, e)
            End Sub

            Public Sub [RaiseEvent](sender As Object, e As ValueChangedEventArgs) Implements IRefreshBehavior.RaiseEvent
                CType(events(NameOf(ValueChanged)), EventHandler(Of ValueChangedEventArgs)).Invoke(sender, e)
            End Sub

            Public Sub [AddHandler](value As PropertyChangedEventHandler) Implements IRefreshBehavior.AddHandler
                events.AddHandler(NameOf(PropertyChanged), value)
            End Sub

            Public Sub [AddHandler](value As EventHandler(Of ValueChangedEventArgs)) Implements IRefreshBehavior.AddHandler
                events.AddHandler(NameOf(ValueChanged), value)
            End Sub

            Public Sub [RemoveHandler](value As PropertyChangedEventHandler) Implements IRefreshBehavior.RemoveHandler
                events.RemoveHandler(NameOf(PropertyChanged), value)
            End Sub

            Public Sub [RemoveHandler](value As EventHandler(Of ValueChangedEventArgs)) Implements IRefreshBehavior.RemoveHandler
                events.RemoveHandler(NameOf(ValueChanged), value)
            End Sub

            Public Sub SuspendRefreshing() Implements IRefreshBehavior.SuspendRefreshing
                parent.refreshState = New RefreshDisabled(parent, events)
            End Sub

            Public Sub EnableRefreshing() Implements IRefreshBehavior.EnableRefreshing

            End Sub
        End Class

#End Region


    End Class

End Namespace