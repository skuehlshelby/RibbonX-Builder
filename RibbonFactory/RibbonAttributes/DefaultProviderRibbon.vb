Namespace RibbonAttributes

    Friend NotInheritable Class DefaultProviderRibbon
        Implements IDefaultProvider

        Public Function GetDefaults() As AttributeGroup Implements IDefaultProvider.GetDefaults
            Dim defaults As AttributeGroupBuilder = New AttributeGroupBuilder()

            defaults.AddStartFromScratch(startFromScratch:= False)

            Return defaults.Build()
        End Function

    End Class

End NameSpace