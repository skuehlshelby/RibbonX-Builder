Imports RibbonFactory.ComponentInterfaces
Imports RibbonFactory.Enums

Namespace RibbonAttributes

    Friend NotInheritable Class DefaultProvider(Of T As RibbonElement)
        Implements IDefaultProvider

        Public Function GetDefaults() as AttributeGroup Implements IDefaultProvider.GetDefaults
            Dim defaults As AttributeGroupBuilder = New AttributeGroupBuilder()

            defaults.AddID(IdManager.GetID(Of T)())

            For Each interfaceType As Type In GetType(T).GetInterfaces
                Select Case interfaceType
                    Case GetType(IEnabled)
                        defaults.AddEnabled(True)
                    Case GetType(IVisible)
                        defaults.AddVisible(True)
                    Case GetType(ILabel)
                        defaults.AddLabel(String.Empty)
                    Case GetType(IShowLabel)
                        defaults.AddShowLabel(False)
                    Case GetType(IScreenTip)
                        defaults.AddScreentip(String.Empty)
                    Case GetType(ISuperTip)
                        defaults.AddSupertip(String.Empty)
                    Case GetType(ISize)
                        defaults.AddSize(ControlSize.large)
                    Case GetType(IDescription)
                        defaults.AddDescription(String.Empty)
                    Case GetType(IImage)
                        defaults.AddImageMso(ImageMSO.Common.HappyFace)
                    Case GetType(IShowImage)
                        defaults.AddShowImage(False)
                End Select
            Next interfaceType

            Return defaults.Build()
        End Function

    End Class

End NameSpace