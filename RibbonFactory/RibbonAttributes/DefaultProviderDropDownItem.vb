Namespace RibbonAttributes

    Friend NotInheritable Class DefaultProviderDropDownItem
        Implements IDefaultProvider

        Public Function GetDefaults() As AttributeGroup Implements IDefaultProvider.GetDefaults
            Dim defaults As AttributeGroupBuilder = New AttributeGroupBuilder()

            With defaults
                '.AddID()
                .AddLabel(String.Empty)
                .AddScreentip(String.Empty)
                .AddSupertip(String.Empty)
            End With

            Return defaults.Build()
        End Function

    End Class

End NameSpace