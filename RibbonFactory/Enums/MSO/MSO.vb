Namespace Enums.MSO

    Public MustInherit Class MSO
        Protected Sub New(name As String, label As String, controlType As MsoControlType)
            Me.Name = name
            Me.Label = label
            Me.ControlType = controlType
        End Sub
        
        Public ReadOnly Property Name As String
        
        Public ReadOnly Property Label As String
        
        Public ReadOnly Property ControlType As MsoControlType
        
        Public Overrides Function ToString() As String
            Return Name
        End Function
        
        Public Class MsoControlType
            Public Shared ReadOnly Property Button As MsoControlType = New MsoControlType()

            Public Shared ReadOnly Property CheckBox As MsoControlType = New MsoControlType()

            Public Shared ReadOnly Property ComboBox As MsoControlType = New MsoControlType()

            Public Shared ReadOnly Property DropDown As MsoControlType = New MsoControlType()

            Public Shared ReadOnly Property Gallery As MsoControlType = New MsoControlType()

            Public Shared ReadOnly Property Group As MsoControlType = New MsoControlType()
            
            Public Shared ReadOnly Property Menu As MsoControlType = New MsoControlType()
            
            Public Shared ReadOnly Property SplitButton As MsoControlType = New MsoControlType()
            
            Public Shared ReadOnly Property Tab As MsoControlType = New MsoControlType()
            
            Public Shared ReadOnly Property ToggleButton As MsoControlType = New MsoControlType()
            
        End Class
    End Class

End Namespace
