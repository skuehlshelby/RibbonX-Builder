Imports RibbonX.Api
Imports RibbonX.Controls.Attributes
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.Properties

Namespace Controls
    Friend NotInheritable Class LabelControl
        Inherits RibbonElement
        Implements ILabelControl
        Implements IAttributeSource

        Private ReadOnly _attributes As AttributeSet

        Public Sub New(attributes As AttributeSet, Optional tag As Object = Nothing)
            MyBase.New(attributes, tag)
        End Sub

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return _attributes.Get(Of Boolean)(AttributeCategory.Enabled)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.Enabled)
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.Get(Of Boolean)(AttributeCategory.Visibility)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.Visibility)
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.Get(Of String)(AttributeCategory.Label)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.Label)
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return _attributes.Get(Of Boolean)(AttributeCategory.LabelVisibility)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.LabelVisibility)
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return _attributes.Get(Of String)(AttributeCategory.ScreenTip)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.ScreenTip)
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return _attributes.Get(Of String)(AttributeCategory.SuperTip)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.SuperTip)
            End Set
        End Property

        Private Function GetAttributes() As AttributeSet Implements IAttributeSource.GetAttributes
            Return _attributes.Clone()
        End Function

        Public Overrides Function Clone() As Object
            Return New LabelControl(Attributes.Clone(), Tag)
        End Function
    End Class

End Namespace