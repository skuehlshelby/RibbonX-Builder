Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes
Imports stdole

Namespace Containers

    Public Class Group
        Inherits RibbonElement
        Implements ILabel
        Implements IVisible
        Implements IKeyTip
        Implements IImage
        Implements IScreenTip
        Implements ISuperTip
        Implements IEnumerable(Of RibbonElement)

        Private ReadOnly _attributes As AttributeGroup
        Private ReadOnly _controls As List(Of RibbonElement) = New List(Of RibbonElement)()

        Friend Sub New(attributes As AttributeGroup, Optional tag As Object = Nothing)
            MyBase.New(tag)
            _attributes = attributes
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Id).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Join(Environment.NewLine, $"<group { _attributes }>", String.Join(Environment.NewLine, Controls), $"</group>")
            End Get
        End Property

        Public ReadOnly Property Controls As IReadOnlyList(Of RibbonElement)
            Get
                Return _controls
            End Get
        End Property

        Public Function AddControl(control As RibbonElement) As RibbonElement
            _controls.Add(control)
            Return control
        End Function

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Label).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of String)(AttributeName.GetLabel).SetValue(Value)
                
            End Set
        End Property

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

        Public Property Image As IPictureDisp Implements IImage.Image
            Get
                Return _attributes.ReadOnlyLookup(Of IPictureDisp)(AttributeName.GetImage).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of IPictureDisp)(AttributeName.GetImage).SetValue(Value)
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

        Public Function GetEnumerator() As IEnumerator(Of RibbonElement) Implements IEnumerable(Of RibbonElement).GetEnumerator
            Return Controls.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return DirectCast(Controls, IEnumerable).GetEnumerator()
        End Function

    End Class
End Namespace