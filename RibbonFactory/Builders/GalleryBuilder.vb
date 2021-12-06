Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports RibbonFactory.RibbonAttributes

Namespace Builders
    Public NotInheritable Class GalleryBuilder
        Private ReadOnly _builder As ControlBuilder

        Sub New()
            Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of Gallery)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = new ControlBuilder(attributeGroupBuilder)
        End Sub
    End Class
End NameSpace