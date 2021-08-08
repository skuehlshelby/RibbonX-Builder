Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes

Namespace Builders

    Public NotInheritable Class ButtonGroupBuilder
        Implements IInsert(Of ButtonGroupBuilder)
        Implements IVisible(Of ButtonGroupBuilder)

        Private ReadOnly _builder As ControlBuilder
        Private ReadOnly _controls As IList(Of RibbonElement)

        Friend Sub New()
            Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of ButtonGroup)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = new ControlBuilder(attributeGroupBuilder)
            _controls = new List(Of RibbonElement)
        End Sub

        Public Function Build(Optional tag As Object = Nothing) As ButtonGroup
            Return New ButtonGroup(_controls, _builder.Build(), tag:=tag)
        End Function

        Public Function Visible() As ButtonGroupBuilder Implements IVisible(Of ButtonGroupBuilder).Visible
            _builder.Visible()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As ButtonGroupBuilder Implements IVisible(Of ButtonGroupBuilder).Visible
            _builder.Visible(callback)
            Return Me
        End Function

        Public Function Invisible() As ButtonGroupBuilder Implements IVisible(Of ButtonGroupBuilder).Invisible
            _builder.Invisible()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As ButtonGroupBuilder Implements IVisible(Of ButtonGroupBuilder).Invisible
            _builder.Invisible(callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As ButtonGroupBuilder Implements IInsert(Of ButtonGroupBuilder).InsertBefore
            _builder.InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As ButtonGroupBuilder Implements IInsert(Of ButtonGroupBuilder).InsertBefore
            _builder.InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As ButtonGroupBuilder Implements IInsert(Of ButtonGroupBuilder).InsertAfter
            _builder.InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As ButtonGroupBuilder Implements IInsert(Of ButtonGroupBuilder).InsertAfter
            _builder.InsertAfter(qualifiedControl)
            Return Me
        End Function
    End Class
End Namespace