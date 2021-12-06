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

            Dim flattened As ICollection(Of RibbonElement) = New List(Of RibbonElement)
            Flatten(_tabs, flattened)

            Return New Ribbon(flattened, ribbonX)
        End Function

        Private Shared Sub Flatten(elements As IEnumerable(Of RibbonElement), flattened As ICollection(Of RibbonElement))
            For Each element As RibbonElement In elements
                flattened.Add(element)


                Select Case element.GetType()
                    Case GetType(DropDown)
                        Dim dropDown As DropDown = DirectCast(element, DropDown)
                        Flatten(dropDown, flattened)
                        Flatten(dropDown.Buttons, flattened)
                    Case GetType(Gallery)
                        Dim gallery As Gallery = DirectCast(element, Gallery)
                        Flatten(gallery, flattened)
                        Flatten(gallery.Buttons, flattened)
                    Case Else
                        Dim ribbonElements As IEnumerable(Of RibbonElement) = TryCast(element, IEnumerable(Of RibbonElement))

                        If ribbonElements IsNot Nothing Then
                            Flatten(ribbonElements, flattened)
                        End If
                End Select
            Next
        End Sub

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