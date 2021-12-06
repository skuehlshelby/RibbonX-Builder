Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes

Namespace Builders

    Public NotInheritable Class TabBuilder
        Implements IInsert(Of TabBuilder)
        Implements IKeyTip(Of TabBuilder)
        Implements ILabelOnly(Of TabBuilder)
        Implements IVisible(Of TabBuilder)

        Private ReadOnly _builder As ControlBuilder
        Private ReadOnly _groups As ICollection(Of Group)

        Public Sub New()
            Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of Tab)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = new ControlBuilder(attributeGroupBuilder)
            _groups = New List(Of Group)
        End Sub

        Public Function Build(Optional tag As Object = Nothing) As Tab
            Return New Tab(_groups, _builder.Build(), tag:=tag)
        End Function

        Public Function Visible() As TabBuilder Implements IVisible(Of TabBuilder).Visible
            _builder.Visible()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As TabBuilder Implements IVisible(Of TabBuilder).Visible
            _builder.Visible(callback)
            Return Me
        End Function

        Public Function Invisible() As TabBuilder Implements IVisible(Of TabBuilder).Invisible
            _builder.Invisible()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As TabBuilder Implements IVisible(Of TabBuilder).Invisible
            _builder.Invisible(callback)
            Return Me
        End Function

        Public Function WithLabel(label As String) As TabBuilder Implements ILabelOnly(Of TabBuilder).WithLabel
            _builder.WithLabel(label:= label)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As TabBuilder Implements IKeyTip(Of TabBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As TabBuilder Implements IKeyTip(Of TabBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip, callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As TabBuilder Implements IInsert(Of TabBuilder).InsertBefore
            _builder.InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As TabBuilder Implements IInsert(Of TabBuilder).InsertBefore
            _builder.InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As TabBuilder Implements IInsert(Of TabBuilder).InsertAfter
            _builder.InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As TabBuilder Implements IInsert(Of TabBuilder).InsertAfter
            _builder.InsertAfter(qualifiedControl)
            Return Me
        End Function

        Public Function WithGroup(group As Group) As TabBuilder
            _groups.Add(group)
            Return Me
        End Function

        Public Function WithGroups(ParamArray groups As Group()) As TabBuilder
            For Each group As Group In groups
                _groups.Add(group)
            Next

            Return Me
        End Function

    End Class

End NameSpace