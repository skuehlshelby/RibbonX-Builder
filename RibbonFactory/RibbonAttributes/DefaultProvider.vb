
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.Enums

Namespace RibbonAttributes

    Friend NotInheritable Class DefaultProvider(Of T As RibbonElement)
        Implements IDefaultProvider

        Public Function GetDefaults() as AttributeSet Implements IDefaultProvider.GetDefaults
            Dim defaults As AttributeSet = New AttributeSet()

            defaults.Add(New RibbonAttributeReadOnly(Of String)(IdManager.GetID(Of T)(), AttributeName.Id, AttributeCategory.IdType))

            'TODO Make sure that everything that should have a default value is represented here
            For Each interfaceType As Type In GetType(T).GetInterfaces
                Select Case interfaceType
                    Case GetType(IEnabled)
                        defaults.Add(New RibbonAttributeDefault(Of Boolean)(True, AttributeName.Enabled, AttributeCategory.Enabled))
                    Case GetType(IVisible)
                        defaults.Add(New RibbonAttributeDefault(Of Boolean)(True, AttributeName.Visible, AttributeCategory.Visibility))
                    Case GetType(ILabel)
                        defaults.Add(New RibbonAttributeDefault(Of String)(String.Empty, AttributeName.Label, AttributeCategory.Label))
                    Case GetType(IShowLabel)
                        defaults.Add(New RibbonAttributeDefault(Of Boolean)(False, AttributeName.ShowLabel, AttributeCategory.LabelVisibility))
                    Case GetType(IScreenTip)
                        defaults.Add(New RibbonAttributeDefault(Of String)(String.Empty, AttributeName.Screentip, AttributeCategory.ScreenTip))
                    Case GetType(ISuperTip)
                        defaults.Add(New RibbonAttributeDefault(Of String)(String.Empty, AttributeName.Supertip, AttributeCategory.SuperTip))
                    Case GetType(ISize)
                        defaults.Add(New RibbonAttributeDefault(Of ControlSize)(ControlSize.normal, AttributeName.Size, AttributeCategory.Size))
                    Case GetType(IDescription)
                        defaults.Add(New RibbonAttributeDefault(Of String)(String.Empty, AttributeName.Description, AttributeCategory.Description))
                    Case GetType(IImage)
                        defaults.Add(New RibbonAttributeDefault(Of ImageMSO.ImageMSO)(ImageMSO.Common.HappyFace, AttributeName.Image, AttributeCategory.Image))
                    Case GetType(IShowImage)
                        defaults.Add(New RibbonAttributeDefault(Of Boolean)(False, AttributeName.ShowImage, AttributeCategory.ImageVisibility))
                    Case GetType(ITitle)
                        defaults.Add(New RibbonAttributeDefault(Of String)(String.Empty, AttributeName.Title, AttributeCategory.Title))
                End Select
            Next interfaceType

            Return defaults
        End Function

    End Class

End NameSpace