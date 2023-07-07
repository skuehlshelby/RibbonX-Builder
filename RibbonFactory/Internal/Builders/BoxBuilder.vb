
Imports RibbonX.Api
Imports RibbonX.Api.Internal
Imports RibbonX.Callbacks
Imports RibbonX.Controls

Namespace InternalApi

    Friend Class BoxBuilder
        Inherits BuilderBase(Of Box)
        Implements IBoxBuilder

        Private ReadOnly items As ISet(Of IRibbonElement) = New HashSet(Of IRibbonElement)()

        Public Overrides Function Build() As Box
            Return New Box(ControlProperties, items.ToArray(), Tag)
        End Function

        Public Function WithTag(tag As Object) As IBoxBuilder Implements ITag(Of IBoxBuilder).WithTag
            WithTagBase(tag)
            Return Me
        End Function

        Public Function Horizontal() As IBoxBuilder Implements IBoxStyle(Of IBoxBuilder).Horizontal
            HorizontalBase()
            Return Me
        End Function

        Public Function Vertical() As IBoxBuilder Implements IBoxStyle(Of IBoxBuilder).Vertical
            VerticalBase()
            Return Me
        End Function

        Public Function Visible() As IBoxBuilder Implements IVisible(Of IBoxBuilder).Visible
            VisibleBase()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As IBoxBuilder Implements IVisible(Of IBoxBuilder).Visible
            VisibleBase(callback)
            Return Me
        End Function

        Public Function Invisible() As IBoxBuilder Implements IVisible(Of IBoxBuilder).Invisible
            InvisibleBase()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As IBoxBuilder Implements IVisible(Of IBoxBuilder).Invisible
            InvisibleBase(callback)
            Return Me
        End Function

        Public Function FromTemplate(template As IRibbonElement) As IBoxBuilder Implements ITemplatable(Of IBoxBuilder).FromTemplate
            FromTemplateBase(template)
            Return Me
        End Function

        Public Function WithControls(ParamArray controls() As IBoxAddable) As IBoxBuilder Implements IBoxBuilder.WithControls
            For Each control As IBoxAddable In controls
                items.Add(control)
            Next

            Return Me
        End Function

        Public Shared Function FromAction(Optional action As Action(Of IBoxBuilder) = Nothing) As IBox
            Dim instance As BoxBuilder = New BoxBuilder()

            action?.Invoke(instance)

            Return instance.Build()
        End Function

    End Class

End Namespace