Imports RibbonFactory.BuilderInterfaces.API
Imports RibbonFactory.RibbonAttributes

Namespace Builders

    Friend NotInheritable Class RibbonBuilder
        Implements IRibbonBuilder

        Private ReadOnly _attributes As AttributeSet = New AttributeSet() From {
            New RibbonAttributeDefault(Of Boolean)(False, AttributeName.StartFromScratch, AttributeCategory.StartFromScratch),
            New RibbonAttributeDefault(Of String)(String.Empty, AttributeName.OnLoad, AttributeCategory.OnLoad),
            New RibbonAttributeDefault(Of String)(String.Empty, AttributeName.LoadImage, AttributeCategory.LoadImage)
        }

        Public Function Build() As AttributeSet
            Return _attributes
        End Function

        Private Function StartFromScratch() As IRibbonBuilder Implements IRibbonBuilder.StartFromScratch
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(True, AttributeName.StartFromScratch, AttributeCategory.StartFromScratch))
            Return Me
        End Function

        Private Function OnLoad(callback As OnLoad) As IRibbonBuilder Implements IRibbonBuilder.OnLoad
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.OnLoad, AttributeCategory.OnLoad))
            Return Me
        End Function

        Private Function LoadImagesFrom(callback As LoadImage) As IRibbonBuilder Implements IRibbonBuilder.LoadImagesFrom
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.LoadImage, AttributeCategory.LoadImage))
            Return Me
        End Function

    End Class

End Namespace