Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.Description
Imports RibbonFactory.RibbonAttributes.Categories.Enabled
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

        Friend Sub New(buttonAttributes As AttributeGroup, Optional tag As Object = Nothing)
            MyBase.New(GetDefaults(Of Button)(), tag)
            Attributes.MergeWith(buttonAttributes)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return _
                    String.Format(String.Join(vbNewLine, "<button {0}/>"),
                                  String.Join(vbNewLine & vbTab, Attributes.Select(Function(a) a.XML)))
            End Get
        End Property

        Public Property Description As String Implements IDescribe.Description
            Get
                Return Attributes.Lookup(Of Description).GetValue()
            End Get
            Set
                Attributes.Lookup(Of GetDescription).SetValue(value)
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return Attributes.Lookup(Of Label).GetValue()
            End Get
            Set
                Attributes.Lookup(Of GetLabel).SetValue(value)
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements ILabel.ShowLabel
            Get
                Return Attributes.Lookup(Of ShowLabel).GetValue()
            End Get
            Set
                Attributes.Lookup(Of GetShowLabel).SetValue(value)
            End Set
        End Property

        Public Property ScreenTip As String Implements ITip.ScreenTip
            Get
                Return Attributes.Lookup(Of Screentip).GetValue()
            End Get
            Set
                Attributes.Lookup(Of GetScreentip).SetValue(value)
            End Set
        End Property

        Public Property SuperTip As String Implements ITip.Supertip
            Get
                Return Attributes.Lookup(Of Supertip).GetValue()
            End Get
            Set
                Attributes.Lookup(Of GetSupertip).SetValue(value)
            End Set
        End Property

        Public Property Enabled As Boolean Implements IEnable.Enabled
            Get
                Return Attributes.Lookup(Of Enabled).GetValue()
            End Get
            Set
                Attributes.Lookup(Of GetEnabled).SetValue(value)
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return Attributes.Lookup(Of Visible).GetValue()
            End Get
            Set
                Attributes.Lookup(Of GetVisible).SetValue(value)
            End Set
        End Property

        Public Property Image As Object Implements IGraphic.Image
            Get
                If IsCustom Then
                    Return Attributes.Lookup(Of GetImage).GetValue()
                Else
                    Return Attributes.Lookup(Of Categories.Image.ImageMso).GetValue()
                End If
            End Get
            Set
                If IsCustom Then
                    Attributes.Lookup(Of GetImage).SetValue(CType(value, IPictureDisp))
                Else
                    Attributes.Lookup(Of Categories.Image.ImageMso).GetValue()
                End If
            End Set
        End Property

        Public ReadOnly Property IsCustom As Boolean Implements IGraphic.IsCustom
            Get
                Return TypeOf Attributes.Lookup(Of ImageBase) Is GetImage
            End Get
        End Property

        Public Property ShowImage As Boolean Implements IGraphic.ShowImage
            Get
                Return Attributes.Lookup(Of ShowImage).GetValue()
            End Get
            Set
                Attributes.Lookup(Of GetShowImage).SetValue(value)
            End Set
        End Property

        Public Property Size As ControlSize Implements IResizable.Size
            Get
                Return Attributes.Lookup(Of Size).GetValue()
            End Get
            Set
                Attributes.Lookup(Of GetSize).SetValue(value)
            End Set
        End Property

        Public Sub Execute() Implements IExecute.Execute
            Attributes.Lookup(Of Categories.OnAction.OnAction).GetValue().Invoke()
        End Sub

        Public Overrides Function Equals(obj As Object) As Boolean
            Return obj.GetHashCode() = GetHashCode() AndAlso TypeOf obj Is Button
        End Function
    End Class
End NameSpace