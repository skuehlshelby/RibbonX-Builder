Imports RibbonFactory.ComponentInterfaces
Imports RibbonFactory.RibbonAttributes


Namespace Containers

    Public NotInheritable Class Tab
        Inherits Container(Of Group)
        Implements IVisible
        Implements IKeyTip
        Implements ILabel

        Private ReadOnly _attributes As AttributeSet

        Friend Sub New(groups As ICollection(Of Group), attributes As AttributeSet, Optional tag As Object = Nothing)
            MyBase.New(groups, tag)
            _attributes = attributes
            AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Id).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return _
                    String.Join(Environment.NewLine, $"<tab { _attributes }>",
                                String.Join(Environment.NewLine, Items), "</tab>")
            End Get
        End Property

        Friend Overrides Sub Flatten(results As ICollection(Of RibbonElement))
            results.Add(Me)

            For Each group As Group In Items
                group.Flatten(results)
            Next
        End Sub

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Visible).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetVisible).SetValue(Value)
            End Set
        End Property

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return _attributes.ReadOnlyLookup(Of KeyTip)(AttributeName.Keytip).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of KeyTip)(AttributeName.GetKeytip).SetValue(Value)
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Label).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of String)(AttributeName.GetLabel).SetValue(Value)
            End Set
        End Property

    End Class
End NameSpace