Imports RibbonFactory.Component_Interfaces

Public Class ComboBox
    Inherits RibbonElement
    Implements IEnabled
    Implements IVisible
    Implements ILabel
    Implements IScreenTip
    Implements ISuperTip
    Implements IShowLabel
    Implements IKeyTip
    Implements IImage
    Implements IShowImage
    Implements IMaxLength
    Implements IOnChange
    Implements IText

    Public Overrides ReadOnly Property ID As String
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides ReadOnly Property XML As String
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Property Enabled As Boolean Implements IEnabled.Enabled
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Boolean)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Property Visible As Boolean Implements IVisible.Visible
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Boolean)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Property Label As String Implements ILabel.Label
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As String)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Property ScreenTip As String Implements IScreenTip.ScreenTip
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As String)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Property SuperTip As String Implements ISuperTip.SuperTip
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As String)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Boolean)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As KeyTip)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Property Image As Object Implements IImage.Image
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Object)
            Throw New NotImplementedException()
        End Set
    End Property

    Public ReadOnly Property IsCustom As Boolean Implements IImage.IsCustom
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Property ShowImage As Boolean Implements IShowImage.ShowImage
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Boolean)
            Throw New NotImplementedException()
        End Set
    End Property

    Public ReadOnly Property MaxLength As UShort Implements IMaxLength.MaxLength
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Property Text As String Implements IText.Text
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As String)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Sub Execute() Implements IOnChange.Execute
        Throw New NotImplementedException()
    End Sub

    Public Overrides Function Equals(obj As Object) As Boolean
        Throw New NotImplementedException()
    End Function
End Class