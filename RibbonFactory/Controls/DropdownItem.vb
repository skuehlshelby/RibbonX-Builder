Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes
Imports stdole

Namespace Controls

    Public NotInheritable Class DropdownItem
        Implements ILabel
        Implements IScreenTip
        Implements ISuperTip
        Implements IImage

#Region "Shared Methods"
        'TODO Move this somewhere else
        Private Shared ReadOnly IdsInCirculation As IDictionary(Of UShort, WeakReference(Of DropdownItem)) = New Dictionary(Of UShort, WeakReference(Of DropdownItem))

        Private Shared Function GetIdSuffix() As UShort
            CleanUpUnusedIds()

            For idSuffix As UShort = UShort.MinValue To UShort.MaxValue
                If Not IdsInCirculation.Keys.Contains(idSuffix) Then
                    Return idSuffix
                End If
            Next

            Throw New Exception($"All available {NameOf(DropdownItem)} IDs are currently in use.")
        End Function

        Private Shared Sub CleanUpUnusedIds()
            For Each id As UShort In IdsInCirculation.Keys
                If Not IdsInCirculation.Item(id).TryGetTarget(Nothing) Then 'Does it still reference something?
                    IdsInCirculation.Remove(id)
                End If
            Next id
        End Sub

#End Region

        Private ReadOnly _attributes As AttributeGroup

        Friend Sub New(attributes As AttributeGroup)
            _attributes = attributes
        End Sub

        Public ReadOnly Property ID As String
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Id).GetValue()
            End Get
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Label).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of String)(AttributeName.GetLabel).SetValue(Value)
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Screentip).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of String)(AttributeName.GetScreentip).SetValue(Value)
            End Set
        End Property

        Public Property SuperTip As String Implements ISupertip.Supertip
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Supertip).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of String)(AttributeName.GetSupertip).SetValue(Value)
            End Set
        End Property

        Public Property Image As IPictureDisp Implements IImage.Image
            Get
                Return _attributes.ReadOnlyLookup(Of IPictureDisp)(AttributeName.GetImage).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of IPictureDisp)(AttributeName.GetImage).SetValue(Value)
            End Set
        End Property

    End Class

End NameSpace