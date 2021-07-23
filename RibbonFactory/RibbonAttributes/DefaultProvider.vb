Namespace RibbonAttributes

    Friend NotInheritable Class DefaultProvider(Of T As RibbonElement)
        Implements IDefaultProvider

        Public Function GetDefaults() as AttributeGroup Implements IDefaultProvider.GetDefaults
            Dim defaults As AttributeGroupBuilder = New AttributeGroupBuilder()

            defaults.AddID(IdManager.GetID(Of T)())

            For Each interfaceType As Type In GetType(T).GetInterfaces
                Select Case interfaceType
                    'TODO
                    'Case GetType(IEnabled)
                    '    defaults.AddEnabled(True)
                    'Case GetType(IVisible)
                    '    GetDefaults.Add(New Visible(True))
                    'Case GetType(ILabel)
                    '    GetDefaults.Add(New Label(String.Empty))
                    'Case GetType(IShowLabel)
                    '    GetDefaults.Add(New ShowLabel(False))
                    'Case GetType(IScreenTip)
                    '    GetDefaults.Add(New Screentip(String.Empty))
                    'Case GetType(ISuperTip)
                    '    GetDefaults.Add(New Supertip(String.Empty))
                    'Case GetType(ISize)
                    '    GetDefaults.Add(New Attributes.Size(ControlSize.large))
                    'Case GetType(IDescription)
                    '    GetDefaults.Add(New Description(String.Empty))
                    'Case GetType(IImage)
                    '    GetDefaults.Add(New Categories.Image.ImageMso(Common.HappyFace))
                End Select
            Next interfaceType

            Return defaults.Build()
        End Function

    End Class

End NameSpace