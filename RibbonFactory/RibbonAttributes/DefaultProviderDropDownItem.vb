Imports RibbonFactory.Controls

Namespace RibbonAttributes

    Friend NotInheritable Class DefaultProviderDropDownItem
        Implements IDefaultProvider

        Public Function GetDefaults() As AttributeGroup Implements IDefaultProvider.GetDefaults
            Dim defaults As AttributeGroupBuilder = New AttributeGroupBuilder()

            With defaults
                .AddID(GetId())
                .AddLabel(String.Empty)
                .AddScreentip(String.Empty)
                .AddSupertip(String.Empty)
            End With

            Return defaults.Build()
        End Function

#Region "ID Management"

        Private Shared ReadOnly DropDownItemIDs As IDictionary(Of String, WeakReference(Of DropdownItem)) = New Dictionary(Of String, WeakReference(Of DropdownItem))

        Private Shared Function GetId() As String
            CleanUpUnusedIds()

            For idSuffix As UShort = UShort.MinValue To UShort.MaxValue
                If Not DropDownItemIDs.Keys.Contains($"item{idSuffix}") Then
                    Return $"item{idSuffix}"
                End If
            Next

            Throw New Exception($"All available {NameOf(DropdownItem)} IDs are currently in use.")
        End Function

        Private Shared Sub CleanUpUnusedIds()
            For Each id As String In DropDownItemIDs.Keys
                If Not DropDownItemIDs.Item(id).TryGetTarget(Nothing) Then 'Does it still reference something?
                    DropDownItemIDs.Remove(id)
                End If
            Next id
        End Sub

#End Region

    End Class

End NameSpace