Imports RibbonX.Api
Imports RibbonX.Callbacks
Imports RibbonX.Controls
Imports RibbonX.InternalApi

Namespace Builders

    Friend NotInheritable Class RibbonBuilder
        Implements IRibbonBuilder

        Private ReadOnly tabs As ISet(Of ITab) = New HashSet(Of ITab)()

        Private ReadOnly properties As IPropertyCollection = New PropertyCollection() From {
            RibbonProperty.Create(Name.StartFromScratch, Category.StartFromScratch, False, True)
        }
        Private Function StartFromScratch() As IRibbonBuilder Implements IRibbonBuilder.StartFromScratch
            properties.Add(RibbonProperty.Create(Name.StartFromScratch, Category.StartFromScratch, True))
            Return Me
        End Function

        Private Function OnLoad(callback As OnLoad) As IRibbonBuilder Implements IRibbonBuilder.OnLoad
            properties.Add(RibbonProperty.Create(Name.OnLoad, Category.OnLoad, callback.Method.Name))
            Return Me
        End Function

        Private Function LoadImagesFrom(callback As LoadImage) As IRibbonBuilder Implements IRibbonBuilder.LoadImagesFrom
            properties.Add(RibbonProperty.Create(Name.LoadImage, Category.LoadImage, callback.Method.Name))
            Return Me
        End Function

        Public Function WithTabs(ParamArray tabs() As ITab) As IRibbonBuilder Implements IRibbonBuilder.WithTabs
            For Each tab As ITab In tabs
                Me.tabs.Add(tab)
            Next
            Return Me
        End Function

        Public Function Build() As IRibbon
            Return New Ribbon(properties, tabs.ToArray())
        End Function

        Friend Shared Function FromAction(action As Action(Of IRibbonBuilder)) As IRibbon
            Dim instance As RibbonBuilder = New RibbonBuilder()

            action?.Invoke(instance)

            Return instance.Build()
        End Function

    End Class

End Namespace