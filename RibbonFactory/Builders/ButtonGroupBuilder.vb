Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Enums.MSO

Namespace Builders

    Public Class ButtonGroupBuilder
        Inherits Builder(Of ButtonGroup)
        Implements ISetInsertionPoint(Of ButtonGroupBuilder)
        Implements ISetVisibility(Of ButtonGroupBuilder)

        Public Overrides Function Build(tag As Object) As ButtonGroup
            Return New ButtonGroup(attributes:=Attributes, tag:=tag)
        End Function

        Public Function Visible() As ButtonGroupBuilder Implements ISetVisibility(Of ButtonGroupBuilder).Visible
            AddVisible(visible:=True)
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As ButtonGroupBuilder Implements ISetVisibility(Of ButtonGroupBuilder).Visible
            AddVisible(visible:=True, callback:=callback)
            Return Me
        End Function

        Public Function Invisible() As ButtonGroupBuilder Implements ISetVisibility(Of ButtonGroupBuilder).Invisible
            AddVisible(visible:=False)
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As ButtonGroupBuilder Implements ISetVisibility(Of ButtonGroupBuilder).Invisible
            AddVisible(visible:=False, callback:=callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As ButtonGroupBuilder Implements ISetInsertionPoint(Of ButtonGroupBuilder).InsertBefore
            InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As ButtonGroupBuilder Implements ISetInsertionPoint(Of ButtonGroupBuilder).InsertBefore
            InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As ButtonGroupBuilder Implements ISetInsertionPoint(Of ButtonGroupBuilder).InsertAfter
            InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As ButtonGroupBuilder Implements ISetInsertionPoint(Of ButtonGroupBuilder).InsertAfter
            InsertAfter(qualifiedControl)
            Return Me
        End Function
    End Class
End Namespace