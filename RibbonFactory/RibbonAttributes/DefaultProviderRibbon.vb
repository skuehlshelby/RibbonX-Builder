Namespace RibbonAttributes

    Friend NotInheritable Class DefaultProviderRibbon
        Implements IDefaultProvider

        Public Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
            Dim defaults As AttributeGroupBuilder = New AttributeGroupBuilder()

            With defaults
                .AddStartFromScratch(startFromScratch:= False)
            End With

            Return defaults.Build()
        End Function

    End Class

End NameSpace