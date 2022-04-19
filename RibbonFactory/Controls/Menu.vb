Imports System.Text.RegularExpressions
Imports RibbonFactory.BuilderInterfaces.API
Imports RibbonFactory.Builders
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes

Namespace Controls

    Public NotInheritable Class Menu
        Inherits Container(Of RibbonElement)
        Implements IVisible
        Implements IEnabled
        Implements ILabel
        Implements IShowLabel
        Implements IDescription
        Implements IScreenTip
        Implements ISuperTip
        Implements IKeyTip
        Implements IImage
        Implements IShowImage
        Implements ISize
        Implements IItemSize
        Implements IDefaultProvider

        Private ReadOnly _attributes As AttributeSet

        Public Sub New(configuration As Action(Of IMenuBuilder), items As MenuControls, Optional tag As Object = Nothing)
            Me.New(Nothing, configuration, items, tag)
        End Sub

        Public Sub New(template As RibbonElement, configuration As Action(Of IMenuBuilder), items As MenuControls, Optional tag As Object = Nothing)
            MyBase.New(If(items, New MenuControls()).ToArray(), tag)

            Dim builder As MenuBuilder = If(template Is Nothing, New MenuBuilder(), New MenuBuilder(template))

            configuration.Invoke(builder)

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
                Dim regex As Regex = New Regex("(?:size|getSize)=""\w+""")

                Return String.Join(Environment.NewLine, $"<menu { _attributes }>", regex.Replace(String.Join(Environment.NewLine, Items), String.Empty), "</menu>")
            End Get
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.Read(Of Boolean)(AttributeCategory.Visibility)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Visibility)
            End Set
        End Property

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return _attributes.Read(Of Boolean)(AttributeCategory.Enabled)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Enabled)
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

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return _attributes.Read(Of String)(AttributeCategory.ScreenTip)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.ScreenTip)
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return _attributes.Read(Of String)(AttributeCategory.SuperTip)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.SuperTip)
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

        Public Property Image As RibbonImage Implements IImage.Image
            Get
                Return _attributes.Read(Of RibbonImage)()
            End Get
            Set
                _attributes.Write(Value)
            End Set
        End Property

        Public Property ShowImage As Boolean Implements IShowImage.ShowImage
            Get
                Return _attributes.Read(Of Boolean)(AttributeCategory.ImageVisibility)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.ImageVisibility)
            End Set
        End Property

        Public Property Description As String Implements IDescription.Description
            Get
                Return _attributes.Read(Of String)(AttributeCategory.Description)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Description)
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return _attributes.Read(Of Boolean)(AttributeCategory.LabelVisibility)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.LabelVisibility)
            End Set
        End Property

        Public ReadOnly Property ItemSize As ControlSize Implements IItemSize.ItemSize
            Get
                Return _attributes.Read(Of ControlSize)(AttributeCategory.ItemSize)
            End Get
        End Property

        Public Property Size As ControlSize Implements ISize.Size
            Get
                Return _attributes.Read(Of ControlSize)(AttributeCategory.Size)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Size)
            End Set
        End Property

        Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
            Return _attributes.Clone()
        End Function

    End Class

End Namespace