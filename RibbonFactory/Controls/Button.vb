Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.Description
Imports RibbonFactory.RibbonAttributes.Categories.Enabled
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.Image
Imports RibbonFactory.RibbonAttributes.Categories.Label
Imports RibbonFactory.RibbonAttributes.Categories.Screentip
Imports RibbonFactory.RibbonAttributes.Categories.Size
Imports RibbonFactory.RibbonAttributes.Categories.Supertip
Imports RibbonFactory.RibbonAttributes.Categories.Visible
Imports stdole

Namespace Controls

    Public NotInheritable Class Button
        Inherits RibbonElement
        Implements IExecute
        Implements IEnable
        Implements ILabel
        Implements ITip
        Implements IDescribe
        Implements IGraphic
        Implements IResizable
        Implements IVisible

        Private ReadOnly _attributes As AttributeGroup = New AttributeGroup()

        Friend Sub New(buttonAttributes As AttributeGroup, Optional tag As Object = Nothing)
            MyBase.New(tag)
            _attributes.MergeWith(buttonAttributes)
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.Lookup(Of Id).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return _
                    String.Format(String.Join(vbNewLine, "<button {0}/>"),
                                  String.Join(vbNewLine & vbTab, _attributes))
            End Get
        End Property

        Public Property Description As String Implements IDescribe.Description
            Get
                Return _attributes.Lookup(Of Description).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetDescription).SetValue(value)
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.Lookup(Of Label).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetLabel).SetValue(value)
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements ILabel.ShowLabel
            Get
                Return _attributes.Lookup(Of ShowLabel).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetShowLabel).SetValue(value)
            End Set
        End Property

        Public Property ScreenTip As String Implements ITip.ScreenTip
            Get
                Return _attributes.Lookup(Of Screentip).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetScreentip).SetValue(value)
            End Set
        End Property

        Public Property SuperTip As String Implements ITip.Supertip
            Get
                Return _attributes.Lookup(Of Supertip).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetSupertip).SetValue(value)
            End Set
        End Property

        Public Property Enabled As Boolean Implements IEnable.Enabled
            Get
                Return _attributes.Lookup(Of Enabled).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetEnabled).SetValue(value)
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.Lookup(Of Visible).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetVisible).SetValue(value)
            End Set
        End Property

        Public Property Image As Object Implements IGraphic.Image
            Get
                If IsCustom Then
                    Return _attributes.Lookup(Of GetImage).GetValue()
                Else
                    Return _attributes.Lookup(Of Categories.Image.ImageMso).GetValue()
                End If
            End Get
            Set
                If IsCustom Then
                    _attributes.Lookup(Of GetImage).SetValue(CType(value, IPictureDisp))
                Else
                    _attributes.Lookup(Of Categories.Image.ImageMso).GetValue()
                End If
            End Set
        End Property

        Public ReadOnly Property IsCustom As Boolean Implements IGraphic.IsCustom
            Get
                Return TypeOf _attributes.Lookup(Of ImageBase) Is GetImage
            End Get
        End Property

        Public Property ShowImage As Boolean Implements IGraphic.ShowImage
            Get
                Return _attributes.Lookup(Of ShowImage).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetShowImage).SetValue(value)
            End Set
        End Property

        Public Property Size As ControlSize Implements IResizable.Size
            Get
                Return _attributes.Lookup(Of Size).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetSize).SetValue(value)
            End Set
        End Property

        Public Sub Execute() Implements IExecute.Execute
            _attributes.Lookup(Of Categories.OnAction.OnAction).GetValue().Invoke()
        End Sub

        Public Overrides Function Equals(obj As Object) As Boolean
            Return obj.GetHashCode() = GetHashCode() AndAlso TypeOf obj Is Button
        End Function
    End Class
End NameSpace