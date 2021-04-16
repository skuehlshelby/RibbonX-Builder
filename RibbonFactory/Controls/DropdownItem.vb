Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.Image
Imports RibbonFactory.RibbonAttributes.Categories.Label
Imports RibbonFactory.RibbonAttributes.Categories.Screentip
Imports RibbonFactory.RibbonAttributes.Categories.Supertip
Imports stdole

Namespace Controls

    Public Class DropdownItem
        Implements ILabel
        Implements IScreenTip
        Implements ISuperTip
        Implements IImage

#Region "Shared Methods"

        Private Shared ReadOnly IdsInCirculation As Dictionary(Of UShort, WeakReference(Of DropdownItem)) = New Dictionary(Of UShort, WeakReference(Of DropdownItem))

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
                If Not IdsInCirculation.Item(id).TryGetTarget(Nothing) Then
                    IdsInCirculation.Remove(id)
                End If
            Next id
        End Sub

        Private Shared Function GetDefaults() As AttributeGroup
            Return New AttributeGroup() From {
                New Id(NameOf(DropdownItem) & GetIdSuffix()), 
                New Label(String.Empty), 
                New Screentip(String.Empty), 
                New Supertip(String.Empty)
            }
        End Function

#End Region

        Private ReadOnly _attributes As AttributeGroup

        Friend Sub New(attributes As AttributeGroup)
            _attributes = GetDefaults()
            _attributes.MergeWith(attributes)
        End Sub

        Public ReadOnly Property ID As String
            Get
                Return _attributes.Lookup(Of Id).GetValue()
            End Get
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.Lookup(Of Screentip).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetScreentip).SetValue(value)
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return _attributes.Lookup(Of Screentip).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetScreentip).SetValue(value)
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return _attributes.Lookup(Of Screentip).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetScreentip).SetValue(value)
            End Set
        End Property

        Private Property Image As IPictureDisp Implements IImage.Image
            Get
                Return _attributes.Lookup(Of GetImage).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetImage).SetValue(CType(value, IPictureDisp))
            End Set
        End Property

        Private ReadOnly Property IsCustom As Boolean Implements IImage.IsCustom
            Get
                Return TypeOf _attributes.Lookup(Of ImageBase) Is GetImage
            End Get
        End Property
    End Class
End NameSpace