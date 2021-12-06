Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.RibbonAttributes

Namespace Builders

    Public NotInheritable Class BoxBuilder
        Implements IBoxStyle(Of BoxBuilder)
        Implements IVisible(Of BoxBuilder)

        Private ReadOnly _builder As ControlBuilder
        Private ReadOnly _items As ICollection(Of RibbonElement)

        Public Sub New()
            Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of Box)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = new ControlBuilder(attributeGroupBuilder)
            _items = New List(Of RibbonElement)
        End Sub

        Public Function Build(tag As Object) As Box
            Return New Box(_builder.Build(), _items.ToArray(), tag)
        End Function

        Public Function Horizontal() As BoxBuilder Implements IBoxStyle(Of BoxBuilder).Horizontal
            _builder.Horizontal()
            Return Me
        End Function

        Public Function Vertical() As BoxBuilder Implements IBoxStyle(Of BoxBuilder).Vertical
            _builder.Vertical()
            Return Me
        End Function
        
        Public Function Visible() As BoxBuilder Implements IVisible(Of BoxBuilder).Visible
            _builder.Visible()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As BoxBuilder Implements IVisible(Of BoxBuilder).Visible
            _builder.Visible(callback:= callback)
            Return Me
        End Function

        Public Function Invisible() As BoxBuilder Implements IVisible(Of BoxBuilder).Invisible
            _builder.Invisible()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As BoxBuilder Implements IVisible(Of BoxBuilder).Invisible
            _builder.Invisible(callback:= callback)
            Return Me
        End Function
        
        Public Function WithItems(ParamArray items() As RibbonElement) As BoxBuilder
            For Each item As RibbonElement In items
                _items.Add(item)
            Next item

            Return Me
        End Function

        Public Shared Function Horizontal(ParamArray items() As RibbonElement) As Box
            Return New BoxBuilder() _
                .Horizontal() _
                .WithItems(items) _
                .Build(tag:= Nothing)
        End Function
        
        Public Shared Function Vertical(ParamArray items() As RibbonElement) As Box
            Return New BoxBuilder() _
                .Vertical() _
                .WithItems(items) _
                .Build(tag:= Nothing)
        End Function
        
    End Class
    
End NameSpace