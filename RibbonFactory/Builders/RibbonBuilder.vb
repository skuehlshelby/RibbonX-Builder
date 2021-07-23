Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.RibbonAttributes

Namespace Builders

    Public NotInheritable Class RibbonBuilder
        Implements IStartFromScratch(Of RibbonBuilder)
        Implements IOnLoad(Of RibbonBuilder)
        Implements ILoadImage(Of RibbonBuilder)

        Private ReadOnly _builder As ControlBuilder

        Friend Sub New()
            Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of SplitButton)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = new ControlBuilder(attributeGroupBuilder)
        End Sub

        Public Function StartFromScratch() As RibbonBuilder Implements IStartFromScratch(Of RibbonBuilder).StartFromScratch
            _builder.StartFromScratch()
            Return Me
        End Function

        Public Function OnLoad(callback As OnLoad) As RibbonBuilder Implements IOnLoad(Of RibbonBuilder).OnLoad
            Throw New NotImplementedException()
        End Function

        Public Function LoadImagesFrom(callback As LoadImage) As RibbonBuilder Implements ILoadImage(Of RibbonBuilder).LoadImagesFrom
            Throw New NotImplementedException()
        End Function

    End Class

End NameSpace