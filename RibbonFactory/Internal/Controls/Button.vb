
Imports System.ComponentModel
Imports System.Linq.Expressions
Imports RibbonX.Api
Imports RibbonX.BindingFactory
Imports RibbonX.Controls.Base
Imports RibbonX.Images
Imports RibbonX.InternalApi
Imports RibbonX.Properties
Imports RibbonX.SimpleTypes

Namespace Controls

    Friend NotInheritable Class Button
        Inherits RibbonElement
        Implements IButton

        Public Sub New(attributes As IPropertyCollection, Optional tag As Object = Nothing)
            MyBase.New(attributes, tag)
        End Sub

#Region "Events"

        Public Custom Event Clicking As EventHandler(Of CancelableEventArgs) Implements IButton.Clicking
            AddHandler(value As EventHandler(Of CancelableEventArgs))
                Attributes.AddHandler(NameOf(Clicking), value)
            End AddHandler

            RemoveHandler(value As EventHandler(Of CancelableEventArgs))
                Attributes.RemoveHandler(NameOf(Clicking), value)
            End RemoveHandler

            RaiseEvent(sender As Object, e As CancelableEventArgs)
                Attributes.RaiseEvent(NameOf(Clicking), sender, e)
            End RaiseEvent
        End Event

        Public Custom Event Clicked As EventHandler Implements IButton.Clicked
            AddHandler(value As EventHandler)
                Attributes.AddHandler(NameOf(Clicked), value)
            End AddHandler

            RemoveHandler(value As EventHandler)
                Attributes.RemoveHandler(NameOf(Clicked), value)
            End RemoveHandler

            RaiseEvent(sender As Object, e As EventArgs)
                Attributes.RaiseEvent(NameOf(Clicked), sender, e)
            End RaiseEvent
        End Event

#End Region

#Region "Enabled and Visible"

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return Attributes.Get(Category.Enabled)
            End Get
            Set
                Attributes.Set(Value, Category.Enabled)
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return Attributes.Get(Category.Visibility)
            End Get
            Set
                Attributes.Set(Value, Category.Visibility)
            End Set
        End Property

#End Region

#Region "Display Text"

        Public Property Description As String Implements IDescription.Description
            Get
                Return Attributes.Get(Category.Description)
            End Get
            Set
                Attributes.Set(Value, Category.Description)
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return Attributes.Get(Category.Label)
            End Get
            Set
                Attributes.Set(Value, Category.Label)
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return Attributes.Get(Category.LabelVisibility)
            End Get
            Set
                Attributes.Set(Value, Category.LabelVisibility)
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return Attributes.Get(Category.ScreenTip)
            End Get
            Set
                Attributes.Set(Value, Category.ScreenTip)
            End Set
        End Property
        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return Attributes.Get(Category.SuperTip)
            End Get
            Set
                Attributes.Set(Value, Category.SuperTip)
            End Set
        End Property

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return Attributes.Get(Category.KeyTip)
            End Get
            Set
                Attributes.Set(Value, Category.KeyTip)
            End Set
        End Property

#End Region

#Region "Image and ShowImage"

        Public Property Image As RibbonImage Implements IImage.Image
            Get
                Return Attributes.Get(Category.Image)
            End Get
            Set
                Attributes.Set(Value, Category.Image)
            End Set
        End Property

        Public Property ShowImage As Boolean Implements IShowImage.ShowImage
            Get
                Return Attributes.Get(Category.ImageVisibility)
            End Get
            Set
                Attributes.Set(Value, Category.ImageVisibility)
            End Set
        End Property

#End Region

#Region "Size"

        Public Property Size As ControlSize Implements ISize.Size
            Get
                Return Attributes.Get(Category.Size)
            End Get
            Set
                Attributes.Set(Value, Category.Size)
            End Set
        End Property

#End Region

#Region "Action"

        Public Sub Click() Implements IClickable.Click
            Dim e As CancelableEventArgs = New CancelableEventArgs()

            RaiseEvent Clicking(Me, e)

            If Not e.IsCancelled Then
                RaiseEvent Clicked(Me, EventArgs.Empty)
            End If
        End Sub

#End Region

        Public Overrides Function Clone() As Object
            Return New Button(CType(Attributes.Clone(), IPropertyCollection), Tag)
        End Function

        Public Overloads Sub Bind(Of TTarget As Class)(target As TTarget, ParamArray expressions() As Expression(Of Action(Of IButton, TTarget))) Implements IButton.Bind
            For Each expression As Expression(Of Action(Of IButton, TTarget)) In expressions
                BindingFactory.Instance.Create(Me, target, expression)
            Next
        End Sub
    End Class

End Namespace