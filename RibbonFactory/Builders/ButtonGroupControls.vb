Imports RibbonX.Controls

Namespace Builders

    Public Class ButtonGroupControls
        Implements IEnumerable(Of RibbonElement)

        Private ReadOnly items As LinkedList(Of RibbonElement) = New LinkedList(Of RibbonElement)

        Public Function Add(ParamArray controls() As Button) As ButtonGroupControls
            Array.ForEach(controls, Sub(c) items.AddLast(c))
            Return Me
        End Function

        Public Function Add(ParamArray controls() As Gallery) As ButtonGroupControls
            Array.ForEach(controls, Sub(c) items.AddLast(c))
            Return Me
        End Function

        Public Function Add(ParamArray controls() As Menu) As ButtonGroupControls
            Array.ForEach(controls, Sub(c) items.AddLast(c))
            Return Me
        End Function

        Public Function Add(ParamArray controls() As Separator) As ButtonGroupControls
            Array.ForEach(controls, Sub(c) items.AddLast(c))
            Return Me
        End Function

        Public Function Add(ParamArray controls() As SplitButton) As ButtonGroupControls
            Array.ForEach(controls, Sub(c) items.AddLast(c))
            Return Me
        End Function

        Public Function Add(ParamArray controls() As ToggleButton) As ButtonGroupControls
            Array.ForEach(controls, Sub(c) items.AddLast(c))
            Return Me
        End Function

        Public Shared Function [From](ParamArray controls() As Button) As ButtonGroupControls
            Return New ButtonGroupControls().Add(controls)
        End Function

        Public Shared Function [From](ParamArray controls() As Gallery) As ButtonGroupControls
            Return New ButtonGroupControls().Add(controls)
        End Function

        Public Shared Function [From](ParamArray controls() As Menu) As ButtonGroupControls
            Return New ButtonGroupControls().Add(controls)
        End Function

        Public Shared Function [From](ParamArray controls() As Separator) As ButtonGroupControls
            Return New ButtonGroupControls().Add(controls)
        End Function

        Public Shared Function [From](ParamArray controls() As SplitButton) As ButtonGroupControls
            Return New ButtonGroupControls().Add(controls)
        End Function

        Public Shared Function [From](ParamArray controls() As ToggleButton) As ButtonGroupControls
            Return New ButtonGroupControls().Add(controls)
        End Function

        Public Shared Function [From](controls As IEnumerable(Of Button)) As ButtonGroupControls
            Return New ButtonGroupControls().Add(controls.ToArray())
        End Function

        Public Shared Function [From](controls As IEnumerable(Of Gallery)) As ButtonGroupControls
            Return New ButtonGroupControls().Add(controls.ToArray())
        End Function

        Public Shared Function [From](controls As IEnumerable(Of Menu)) As ButtonGroupControls
            Return New ButtonGroupControls().Add(controls.ToArray())
        End Function

        Public Shared Function [From](controls As IEnumerable(Of Separator)) As ButtonGroupControls
            Return New ButtonGroupControls().Add(controls.ToArray())
        End Function

        Public Shared Function [From](controls As IEnumerable(Of SplitButton)) As ButtonGroupControls
            Return New ButtonGroupControls().Add(controls.ToArray())
        End Function

        Public Shared Function [From](controls As IEnumerable(Of ToggleButton)) As ButtonGroupControls
            Return New ButtonGroupControls().Add(controls.ToArray())
        End Function

        Private Function GetEnumerator() As IEnumerator(Of RibbonElement) Implements IEnumerable(Of RibbonElement).GetEnumerator
            Return items.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return items.GetEnumerator()
        End Function

    End Class

End Namespace
