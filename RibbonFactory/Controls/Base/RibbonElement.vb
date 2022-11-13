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
        End Sub

        Public Event ValueChanged As EventHandler(Of ValueChangedEventArgs)

        Public MustOverride ReadOnly Property ID As String

        Public Property Tag As Object

        Public MustOverride ReadOnly Property XML As String

#Region "Refreshing"

        Protected Sub RefreshNeeded()
            If refreshDecrement = 0 Then
                RaiseEvent ValueChanged(Me, New ValueChangedEventArgs(ID))
            End If
        End Sub

        Public Function RefreshSuspension(Optional refreshOnDispose As Boolean = True) As IDisposable
            Return New UpdateBlock(Me, refreshOnDispose)
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
            Private ReadOnly refreshOnDispose As Boolean

            Public Sub New(parent As RibbonElement, refreshOnDispose As Boolean)
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