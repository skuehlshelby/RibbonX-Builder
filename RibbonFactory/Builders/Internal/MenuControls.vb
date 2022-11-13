Imports RibbonX.Controls
Imports RibbonX.Controls.Base

Namespace Builders

    Public Class MenuControls

        Implements IEnumerable(Of RibbonElement)

        Private ReadOnly items As LinkedList(Of RibbonElement) = New LinkedList(Of RibbonElement)()

        Public Function Add(ParamArray controls() As IButton) As MenuControls
            Array.ForEach(controls, Sub(c) items.AddLast(CType(c, RibbonElement)))
            Return Me
        End Function

        Public Function Add(ParamArray controls() As CheckBox) As MenuControls
            Array.ForEach(controls, Sub(c) items.AddLast(c))
            Return Me
        End Function

        Public Function Add(ParamArray controls() As Gallery) As MenuControls
            Array.ForEach(controls, Sub(c) items.AddLast(c))
            Return Me
        End Function

        Public Function Add(ParamArray controls() As Menu) As MenuControls
            Array.ForEach(controls, Sub(c) items.AddLast(c))
            Return Me
        End Function

        Public Function Add(ParamArray controls() As MenuSeparator) As MenuControls
            Array.ForEach(controls, Sub(c) items.AddLast(c))
            Return Me
        End Function

        Public Function Add(ParamArray controls() As SplitButton) As MenuControls
            Array.ForEach(controls, Sub(c) items.AddLast(c))
            Return Me
        End Function

        Public Function Add(ParamArray controls() As ToggleButton) As MenuControls
            Array.ForEach(controls, Sub(c) items.AddLast(c))
            Return Me
        End Function

        Public Shared Function [From](ParamArray controls() As IButton) As MenuControls
            Return New MenuControls().Add(controls)
        End Function

        Public Shared Function [From](ParamArray controls() As CheckBox) As MenuControls
            Return New MenuControls().Add(controls)
        End Function

        Public Shared Function [From](ParamArray controls() As Gallery) As MenuControls
            Return New MenuControls().Add(controls)
        End Function

        Public Shared Function [From](ParamArray controls() As Menu) As MenuControls
            Return New MenuControls().Add(controls)
        End Function

        Public Shared Function [From](ParamArray controls() As MenuSeparator) As MenuControls
            Return New MenuControls().Add(controls)
        End Function

        Public Shared Function [From](ParamArray controls() As SplitButton) As MenuControls
            Return New MenuControls().Add(controls)
        End Function

        Public Shared Function [From](ParamArray controls() As ToggleButton) As MenuControls
            Return New MenuControls().Add(controls)
        End Function

        Private Function GetEnumerator() As IEnumerator(Of RibbonElement) Implements IEnumerable(Of RibbonElement).GetEnumerator
            Return items.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return items.GetEnumerator()
        End Function

    End Class

End Namespace
