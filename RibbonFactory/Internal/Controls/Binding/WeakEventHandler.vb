Friend Class WeakEventHandler(Of TEventSource As Class, TEventArgs As EventArgs)
    Public Delegate Sub UnregisterCallback(eventHandler As EventHandler(Of TEventArgs))
    Private Delegate Sub OpenEventHandler(this As TEventSource, sender As Object, e As TEventArgs)

    Private ReadOnly targetRef As WeakReference
    Private ReadOnly openHandler As OpenEventHandler
    Private ReadOnly handler As EventHandler(Of TEventArgs)
    Private unregister As UnregisterCallback

    Public Sub New(eventHandler As EventHandler(Of TEventArgs), unregisterCallback As UnregisterCallback)
        targetRef = New WeakReference(eventHandler.Target)
        openHandler = DirectCast([Delegate].CreateDelegate(GetType(OpenEventHandler), Nothing, eventHandler.Method), OpenEventHandler)
        handler = AddressOf Invoke
        unregister = unregisterCallback
    End Sub

    Public Sub Invoke(sender As Object, e As TEventArgs)
        Dim target As TEventSource = DirectCast(targetRef.Target, TEventSource)

        If target IsNot Nothing Then
            openHandler.Invoke(target, sender, e)
        ElseIf unregister IsNot Nothing Then
            unregister.Invoke(handler)
            unregister = Nothing
        End If
    End Sub

    Public Shared Widening Operator CType(weakEventHandler As WeakEventHandler(Of TEventSource, TEventArgs)) As EventHandler(Of TEventArgs)
        Return weakEventHandler.handler
    End Operator

End Class
