Imports RibbonFactory.Controls

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

        Private Function GetEnumerator() As IEnumerator(Of RibbonElement) Implements IEnumerable(Of RibbonElement).GetEnumerator
            Return items.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return items.GetEnumerator()
        End Function

    End Class

End Namespace
