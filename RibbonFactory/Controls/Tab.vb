Imports RibbonX.Builders
Imports RibbonX.Controls.Attributes
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.Properties
Imports RibbonX.SimpleTypes

Namespace Controls

    Public NotInheritable Class Tab
        Inherits Container(Of Group)
        Implements IVisible
        Implements IKeyTip
        Implements ILabel
        Implements IAttributeSource

        Private ReadOnly _attributes As AttributeSet

        Public Sub New(ParamArray groups() As Group)
            Me.New(Nothing, groups)
        End Sub

        Public Sub New(configuration As Action(Of ITabBuilder), ParamArray groups() As Group)
            Me.New(configuration, groups, Nothing)
        End Sub

        Public Sub New(configuration As Action(Of ITabBuilder), groups As ICollection(Of Group), template As RibbonElement, Optional tag As Object = Nothing)
            MyBase.New(groups, tag)

            Dim builder As TabBuilder = New TabBuilder(template)

            If configuration IsNot Nothing Then
                configuration.Invoke(builder)
            End If

            _attributes = builder.Build()

            AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.Read(Of String)(AttributeCategory.IdType)
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                If Items.Any() Then
                    Return String.Join(Environment.NewLine, $"<tab { _attributes }>", String.Join(Environment.NewLine, Items), "</tab>")
                Else
                    Return $"<tab { _attributes } />"
                End If
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
                Return _attributes.Read(Of Boolean)(AttributeCategory.Visibility)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Visibility)
            End Set
        End Property

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return _attributes.Read(Of KeyTip)()
            End Get
            Set
                _attributes.Write(Value)
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.Read(Of String)(AttributeCategory.Label)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Label)
            End Set
        End Property

        Private Function GetAttributes() As AttributeSet Implements IAttributeSource.GetAttributes
            Return _attributes.Clone()
        End Function

    End Class

End Namespace