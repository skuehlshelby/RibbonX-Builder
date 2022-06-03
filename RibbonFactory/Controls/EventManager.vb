Imports RibbonX.RibbonAttributes

Namespace Controls

    Friend Class EventManager(Of TEventArgs As EventArgs)

        Private ReadOnly eventName As String

        Public Sub New(eventName As String)
            Me.eventName = eventName
        End Sub

        Public Sub Add(attributes As AttributeSet, handler As EventHandler(Of TEventArgs))
            attributes.Read(Of ICollection(Of EventHandler(Of TEventArgs)))().Add(handler)
        End Sub

        Public Sub Remove(attributes As AttributeSet, handler As EventHandler(Of TEventArgs))
            attributes.Read(Of ICollection(Of EventHandler(Of TEventArgs)))().Remove(handler)
        End Sub

        Public Sub FireEvent(attributes As AttributeSet, sender As Object, e As TEventArgs)
            For Each target As EventHandler(Of TEventArgs) In attributes.Read(Of ICollection(Of EventHandler(Of TEventArgs)))()
                Try
                    target.Invoke(sender, e)
                Catch ex As Exception
                    Debug.WriteLine($"Encountered an exception while invoking '{target.Method.Name}' inside '{eventName}':")
                    Debug.WriteLine(ex.Message)
                End Try
            Next
        End Sub

    End Class

    Friend Class EventManager

        Private ReadOnly eventName As String

        Public Sub New(eventName As String)
            Me.eventName = eventName
        End Sub

        Public Sub Add(attributes As AttributeSet, handler As EventHandler)
            attributes.Read(Of ICollection(Of EventHandler))().Add(handler)
        End Sub

        Public Sub Remove(attributes As AttributeSet, handler As EventHandler)
            attributes.Read(Of ICollection(Of EventHandler))().Remove(handler)
        End Sub

        Public Sub FireEvent(attributes As AttributeSet, sender As Object, e As EventArgs)
            For Each target As EventHandler In attributes.Read(Of ICollection(Of EventHandler))()
                Try
                    target.Invoke(sender, e)
                Catch ex As Exception
                    Debug.WriteLine($"Encountered an exception while invoking '{target.Method.Name}' inside '{eventName}':")
                    Debug.WriteLine(ex.Message)
                End Try
            Next
        End Sub

    End Class

End Namespace