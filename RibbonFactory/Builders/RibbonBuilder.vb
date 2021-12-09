Imports System.Reflection
Imports System.Xml
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports RibbonFactory.RibbonAttributes

Namespace Builders

    Public NotInheritable Class RibbonBuilder
        Implements IStartFromScratch(Of RibbonBuilder)
        Implements IOnLoad(Of RibbonBuilder)
        Implements ILoadImage(Of RibbonBuilder)

        Private ReadOnly _builder As ControlBuilder
        Private ReadOnly _tabs As IList(Of Tab)

        Public Sub New()
            Dim defaultProvider As IDefaultProvider = New DefaultProviderRibbon()
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = new ControlBuilder(attributeGroupBuilder)
            _tabs = New List(Of Tab)
        End Sub

        Public Function Build() As Ribbon
            Dim attributes As AttributeSet = _builder.Build()

            Dim ribbonX As String = 
                    $"<?xml version=""1.0"" encoding=""utf-8"" ?>
                    
                    <customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" {If(attributes.HasAttribute(AttributeName.OnLoad), attributes.ReadOnlyLookup(Of String)(AttributeName.OnLoad), Nothing)} {If(attributes.HasAttribute(AttributeName.LoadImage), attributes.ReadOnlyLookup(Of String)(AttributeName.LoadImage), Nothing)}>
                        <ribbon {If(attributes.HasAttribute(AttributeName.StartFromScratch), attributes.ReadOnlyLookup(Of Boolean)(AttributeName.StartFromScratch), Nothing)}>
                            <tabs>
                                {String.Join(Environment.NewLine, _tabs.Select(Function(tab) tab.XML))}
                            </tabs>
                        </ribbon>
                    </customUI>"

            Return New Ribbon(Flatten().ToArray(), ribbonX)
        End Function

        Private Function Flatten() As ICollection(Of RibbonElement)
            Dim results As ICollection(Of RibbonElement) = New List(Of RibbonElement)

            For Each tab As Tab In _tabs
                tab.Flatten(results)
            Next

            Return results
        End Function

        Public Function StartFromScratch() As RibbonBuilder Implements IStartFromScratch(Of RibbonBuilder).StartFromScratch
            _builder.StartFromScratch()
            Return Me
        End Function

        Public Function OnLoad(callback As OnLoad) As RibbonBuilder Implements IOnLoad(Of RibbonBuilder).OnLoad
            _builder.OnLoad(callback)
            Return Me
        End Function

        Public Function LoadImagesFrom(callback As LoadImage) As RibbonBuilder Implements ILoadImage(Of RibbonBuilder).LoadImagesFrom
            _builder.LoadImage(callback)
            Return Me
        End Function

        Public Function WithTab(tab As Tab) As RibbonBuilder
            _tabs.Add(tab)
            Return Me
        End Function

        Public Function WithTabs(ParamArray tabs As Tab()) As RibbonBuilder
            For Each tab As Tab In tabs
                _tabs.Add(tab)
            Next

            Return Me
        End Function
    End Class

End NameSpace