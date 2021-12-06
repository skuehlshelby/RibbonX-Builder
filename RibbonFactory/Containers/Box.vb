Imports System.Reflection
Imports RibbonFactory.ComponentInterfaces
Imports RibbonFactory.RibbonAttributes

Namespace Containers

    Public NotInheritable Class Box
        Inherits Container(Of RibbonElement)
        Implements IVisible

        Private ReadOnly _attributes As AttributeSet

        Friend Sub New(attributes As AttributeSet, items As ICollection(Of RibbonElement), Optional tag As Object = Nothing)
            MyBase.New(items, tag)
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
                    String.Join(Environment.NewLine, $"<box { _attributes }>",
                                String.Join(Environment.NewLine, Items), $"</box>")
            End Get
        End Property

        Friend Overrides Sub Flatten(results As ICollection(Of RibbonElement))
            Dim genericContainer As Type = GetType(Container(Of )).GetGenericTypeDefinition()

            For Each item As RibbonElement In Items
                Dim itemType As Type = item.GetType()

                If itemType.IsSubclassOf(genericContainer) Then
                    itemType.InvokeMember(NameOf(Flatten), BindingFlags.Default, Nothing, item, New Object(){results})
                Else
                    results.Add(item)
                End If
            Next
        End Sub

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Visible).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.Visible).SetValue(value)
            End Set
        End Property

    End Class
End Namespace