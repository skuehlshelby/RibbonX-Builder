Imports System.ComponentModel
Imports RibbonX.Utilities

Namespace Controls.Base

    Friend NotInheritable Class PropertyChangedHelper
        Public ReadOnly Property EventQueue As Queue(Of PropertyChangedEventArgs) = New Queue(Of PropertyChangedEventArgs)()
        Public ReadOnly Property Handlers As ICollection(Of PropertyChangedEventHandler) = New LinkedList(Of PropertyChangedEventHandler)()

        Public Sub InvokeAllHandlers(sender As Object, e As PropertyChangedEventArgs)
            For Each handler As PropertyChangedEventHandler In Handlers
                handler?.Invoke(sender, e)
            Next
        End Sub

        Public Sub EmptyQueue(sender As Object)
            While EventQueue.Any()
                InvokeAllHandlers(sender, EventQueue.Dequeue())
            End While
        End Sub
    End Class

    Friend NotInheritable Class RefreshNeededHelper
        Public ReadOnly Property EventQueue As Queue(Of RefreshNeededEventArgs) = New Queue(Of RefreshNeededEventArgs)()
        Public ReadOnly Property Handlers As ICollection(Of EventHandler(Of RefreshNeededEventArgs)) = New LinkedList(Of EventHandler(Of RefreshNeededEventArgs))()

        Public Sub InvokeAllHandlers(sender As Object, e As RefreshNeededEventArgs)
            For Each handler As EventHandler(Of RefreshNeededEventArgs) In Handlers
                handler?.Invoke(sender, e)
            Next
        End Sub

        Public Sub EmptyQueue(sender As Object)
            If EventQueue.Any() Then
                InvokeAllHandlers(sender, EventQueue.Dequeue())
            End If

            EventQueue.Clear()
        End Sub
    End Class

    Friend NotInheritable Class RefreshSuspensionTracker
        Private counter As Integer = 0

        Public Event HitZero As EventHandler

        Public ReadOnly Property HasActiveRefreshSuspension As Boolean
            Get
                Return Not counter = 0
            End Get
        End Property

        Public Function AddSuspension() As IDisposable
            Return New Suspension(Me)
        End Function

        Private Sub TriggerLastSuspensionDisposed()
            RaiseEvent HitZero(Me, EventArgs.Empty)
        End Sub

        Private NotInheritable Class Suspension
            Implements IDisposable
            Private parent As RefreshSuspensionTracker

            Public Sub New(parent As RefreshSuspensionTracker)
                Me.parent = NotNull(parent)
                parent.counter -= 1
            End Sub

            Public Sub Dispose() Implements IDisposable.Dispose
                If parent IsNot Nothing Then
                    parent.counter += 1

                    If parent.counter = 0 Then
                        parent.TriggerLastSuspensionDisposed()
                    End If

                    parent = Nothing
                End If
            End Sub

            Protected Overrides Sub Finalize()
                Dispose()
                MyBase.Finalize()
            End Sub
        End Class
    End Class

End Namespace