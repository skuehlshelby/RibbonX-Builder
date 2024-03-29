﻿Namespace Images.BuiltIn

    Public NotInheritable Class Pivot
        Inherits ImageMSO

        Private Sub New(value As Integer, name As String)
            MyBase.New(value, name)
        End Sub

        Public Shared ReadOnly Property AccessFormPivotTable As Pivot = New Pivot(1, NameOf(AccessFormPivotTable))
        Public Shared ReadOnly Property CreateFormPivotChart As Pivot = New Pivot(2, NameOf(CreateFormPivotChart))
        Public Shared ReadOnly Property GroupPivotActions As Pivot = New Pivot(3, NameOf(GroupPivotActions))
        Public Shared ReadOnly Property GroupPivotChartActiveField As Pivot = New Pivot(4, NameOf(GroupPivotChartActiveField))
        Public Shared ReadOnly Property GroupPivotChartActiveFieldAccess As Pivot = New Pivot(5, NameOf(GroupPivotChartActiveFieldAccess))
        Public Shared ReadOnly Property GroupPivotChartData As Pivot = New Pivot(6, NameOf(GroupPivotChartData))
        Public Shared ReadOnly Property GroupPivotChartDataAccess As Pivot = New Pivot(7, NameOf(GroupPivotChartDataAccess))
        Public Shared ReadOnly Property GroupPivotChartFilterAndSort As Pivot = New Pivot(8, NameOf(GroupPivotChartFilterAndSort))
        Public Shared ReadOnly Property GroupPivotChartShowHide As Pivot = New Pivot(9, NameOf(GroupPivotChartShowHide))
        Public Shared ReadOnly Property GroupPivotChartShowOrHide As Pivot = New Pivot(10, NameOf(GroupPivotChartShowOrHide))
        Public Shared ReadOnly Property GroupPivotChartTools As Pivot = New Pivot(11, NameOf(GroupPivotChartTools))
        Public Shared ReadOnly Property GroupPivotChartType As Pivot = New Pivot(12, NameOf(GroupPivotChartType))
        Public Shared ReadOnly Property GroupPivotTableActiveField As Pivot = New Pivot(13, NameOf(GroupPivotTableActiveField))
        Public Shared ReadOnly Property GroupPivotTableActiveFieldAccess As Pivot = New Pivot(14, NameOf(GroupPivotTableActiveFieldAccess))
        Public Shared ReadOnly Property GroupPivotTableData As Pivot = New Pivot(15, NameOf(GroupPivotTableData))
        Public Shared ReadOnly Property GroupPivotTableDataAccess As Pivot = New Pivot(16, NameOf(GroupPivotTableDataAccess))
        Public Shared ReadOnly Property GroupPivotTableFilterAndSort As Pivot = New Pivot(17, NameOf(GroupPivotTableFilterAndSort))
        Public Shared ReadOnly Property GroupPivotTableGroup As Pivot = New Pivot(18, NameOf(GroupPivotTableGroup))
        Public Shared ReadOnly Property GroupPivotTableLayout As Pivot = New Pivot(19, NameOf(GroupPivotTableLayout))
        Public Shared ReadOnly Property GroupPivotTableOptions As Pivot = New Pivot(20, NameOf(GroupPivotTableOptions))
        Public Shared ReadOnly Property GroupPivotTableSelections As Pivot = New Pivot(21, NameOf(GroupPivotTableSelections))
        Public Shared ReadOnly Property GroupPivotTableShowHide As Pivot = New Pivot(22, NameOf(GroupPivotTableShowHide))
        Public Shared ReadOnly Property GroupPivotTableShowHideAccess As Pivot = New Pivot(23, NameOf(GroupPivotTableShowHideAccess))
        Public Shared ReadOnly Property GroupPivotTableSort As Pivot = New Pivot(24, NameOf(GroupPivotTableSort))
        Public Shared ReadOnly Property GroupPivotTableStyleOptions As Pivot = New Pivot(25, NameOf(GroupPivotTableStyleOptions))
        Public Shared ReadOnly Property GroupPivotTableStyles As Pivot = New Pivot(26, NameOf(GroupPivotTableStyles))
        Public Shared ReadOnly Property GroupPivotTableTools As Pivot = New Pivot(27, NameOf(GroupPivotTableTools))
        Public Shared ReadOnly Property GroupPivotTableToolsAccess As Pivot = New Pivot(28, NameOf(GroupPivotTableToolsAccess))
        Public Shared ReadOnly Property NewPivotTableStyle2 As Pivot = New Pivot(29, NameOf(NewPivotTableStyle2))
        Public Shared ReadOnly Property PivotAutoCalcMenu As Pivot = New Pivot(30, NameOf(PivotAutoCalcMenu))
        Public Shared ReadOnly Property PivotAutoFilter As Pivot = New Pivot(31, NameOf(PivotAutoFilter))
        Public Shared ReadOnly Property PivotChartClearMenu As Pivot = New Pivot(32, NameOf(PivotChartClearMenu))
        Public Shared ReadOnly Property PivotChartFilterShow As Pivot = New Pivot(33, NameOf(PivotChartFilterShow))
        Public Shared ReadOnly Property PivotChartInsert As Pivot = New Pivot(34, NameOf(PivotChartInsert))
        Public Shared ReadOnly Property PivotChartInsertClassic As Pivot = New Pivot(35, NameOf(PivotChartInsertClassic))
        Public Shared ReadOnly Property PivotChartLegendShowHide As Pivot = New Pivot(36, NameOf(PivotChartLegendShowHide))
        Public Shared ReadOnly Property PivotChartMultiplePlots As Pivot = New Pivot(37, NameOf(PivotChartMultiplePlots))
        Public Shared ReadOnly Property PivotChartMultipleUnified As Pivot = New Pivot(38, NameOf(PivotChartMultipleUnified))
        Public Shared ReadOnly Property PivotChartRefresh As Pivot = New Pivot(39, NameOf(PivotChartRefresh))
        Public Shared ReadOnly Property PivotChartSortByTotalMenu As Pivot = New Pivot(40, NameOf(PivotChartSortByTotalMenu))
        Public Shared ReadOnly Property PivotChartType As Pivot = New Pivot(41, NameOf(PivotChartType))
        Public Shared ReadOnly Property PivotClearAll As Pivot = New Pivot(42, NameOf(PivotClearAll))
        Public Shared ReadOnly Property PivotClearCustomOrdering As Pivot = New Pivot(43, NameOf(PivotClearCustomOrdering))
        Public Shared ReadOnly Property PivotCollapseField As Pivot = New Pivot(44, NameOf(PivotCollapseField))
        Public Shared ReadOnly Property PivotCollapseFieldAccess As Pivot = New Pivot(45, NameOf(PivotCollapseFieldAccess))
        Public Shared ReadOnly Property PivotDropAreas As Pivot = New Pivot(46, NameOf(PivotDropAreas))
        Public Shared ReadOnly Property PivotExpandField As Pivot = New Pivot(47, NameOf(PivotExpandField))
        Public Shared ReadOnly Property PivotExpandIndicators As Pivot = New Pivot(48, NameOf(PivotExpandIndicators))
        Public Shared ReadOnly Property PivotExportToExcel As Pivot = New Pivot(49, NameOf(PivotExportToExcel))
        Public Shared ReadOnly Property PivotFieldList As Pivot = New Pivot(50, NameOf(PivotFieldList))
        Public Shared ReadOnly Property PivotFieldListShowHide As Pivot = New Pivot(51, NameOf(PivotFieldListShowHide))
        Public Shared ReadOnly Property PivotFilterBySelection As Pivot = New Pivot(52, NameOf(PivotFilterBySelection))
        Public Shared ReadOnly Property PivotFormulasMenu As Pivot = New Pivot(53, NameOf(PivotFormulasMenu))
        Public Shared ReadOnly Property PivotGroupItems As Pivot = New Pivot(54, NameOf(PivotGroupItems))
        Public Shared ReadOnly Property PivotHideDetails As Pivot = New Pivot(55, NameOf(PivotHideDetails))
        Public Shared ReadOnly Property PivotMoveField As Pivot = New Pivot(56, NameOf(PivotMoveField))
        Public Shared ReadOnly Property PivotMoveToColumnArea As Pivot = New Pivot(57, NameOf(PivotMoveToColumnArea))
        Public Shared ReadOnly Property PivotMoveToDetailArea As Pivot = New Pivot(58, NameOf(PivotMoveToDetailArea))
        Public Shared ReadOnly Property PivotMoveToFieldArea As Pivot = New Pivot(59, NameOf(PivotMoveToFieldArea))
        Public Shared ReadOnly Property PivotMoveToFilterArea As Pivot = New Pivot(60, NameOf(PivotMoveToFilterArea))
        Public Shared ReadOnly Property PivotPlusMinusButtonsShowHide As Pivot = New Pivot(61, NameOf(PivotPlusMinusButtonsShowHide))
        Public Shared ReadOnly Property PivotPlusMinusFieldHeadersShowHide As Pivot = New Pivot(62, NameOf(PivotPlusMinusFieldHeadersShowHide))
        Public Shared ReadOnly Property PivotRefresh As Pivot = New Pivot(63, NameOf(PivotRefresh))
        Public Shared ReadOnly Property PivotRemoveField As Pivot = New Pivot(64, NameOf(PivotRemoveField))
        Public Shared ReadOnly Property PivotShowAsMenu As Pivot = New Pivot(65, NameOf(PivotShowAsMenu))
        Public Shared ReadOnly Property PivotShowDetails As Pivot = New Pivot(66, NameOf(PivotShowDetails))
        Public Shared ReadOnly Property PivotShowOnlyTheBottomMenu As Pivot = New Pivot(67, NameOf(PivotShowOnlyTheBottomMenu))
        Public Shared ReadOnly Property PivotShowOnlyTheTopMenu As Pivot = New Pivot(68, NameOf(PivotShowOnlyTheTopMenu))
        Public Shared ReadOnly Property PivotShowTopAndBottomItemsMenu As Pivot = New Pivot(69, NameOf(PivotShowTopAndBottomItemsMenu))
        Public Shared ReadOnly Property PivotSubtotal As Pivot = New Pivot(70, NameOf(PivotSubtotal))
        Public Shared ReadOnly Property PivotSwitchRowColumn As Pivot = New Pivot(71, NameOf(PivotSwitchRowColumn))
        Public Shared ReadOnly Property PivotTableBlankRowsInsert As Pivot = New Pivot(72, NameOf(PivotTableBlankRowsInsert))
        Public Shared ReadOnly Property PivotTableBlankRowsRemove As Pivot = New Pivot(73, NameOf(PivotTableBlankRowsRemove))
        Public Shared ReadOnly Property PivotTableChangeDataSource As Pivot = New Pivot(74, NameOf(PivotTableChangeDataSource))
        Public Shared ReadOnly Property PivotTableClearMenu As Pivot = New Pivot(75, NameOf(PivotTableClearMenu))
        Public Shared ReadOnly Property PivotTableEditDataSource As Pivot = New Pivot(76, NameOf(PivotTableEditDataSource))
        Public Shared ReadOnly Property PivotTableEnableSelection As Pivot = New Pivot(77, NameOf(PivotTableEnableSelection))
        Public Shared ReadOnly Property PivotTableExpandField As Pivot = New Pivot(78, NameOf(PivotTableExpandField))
        Public Shared ReadOnly Property PivotTableFieldSettings As Pivot = New Pivot(79, NameOf(PivotTableFieldSettings))
        Public Shared ReadOnly Property PivotTableFormulasMenu As Pivot = New Pivot(80, NameOf(PivotTableFormulasMenu))
        Public Shared ReadOnly Property PivotTableGenerateGetPivotData As Pivot = New Pivot(81, NameOf(PivotTableGenerateGetPivotData))
        Public Shared ReadOnly Property PivotTableGrandTotalsOffForRowsAndColumns As Pivot = New Pivot(82, NameOf(PivotTableGrandTotalsOffForRowsAndColumns))
        Public Shared ReadOnly Property PivotTableGrandTotalsOnForColumnsOnly As Pivot = New Pivot(83, NameOf(PivotTableGrandTotalsOnForColumnsOnly))
        Public Shared ReadOnly Property PivotTableGrandTotalsOnForRowsAndColumns As Pivot = New Pivot(84, NameOf(PivotTableGrandTotalsOnForRowsAndColumns))
        Public Shared ReadOnly Property PivotTableGrandTotalsOnForRowsOnly As Pivot = New Pivot(85, NameOf(PivotTableGrandTotalsOnForRowsOnly))
        Public Shared ReadOnly Property PivotTableGroupField As Pivot = New Pivot(86, NameOf(PivotTableGroupField))
        Public Shared ReadOnly Property PivotTableGroupSelection As Pivot = New Pivot(87, NameOf(PivotTableGroupSelection))
        Public Shared ReadOnly Property PivotTableInsert As Pivot = New Pivot(88, NameOf(PivotTableInsert))
        Public Shared ReadOnly Property PivotTableInsertMenu As Pivot = New Pivot(89, NameOf(PivotTableInsertMenu))
        Public Shared ReadOnly Property PivotTableLayoutBlankRows As Pivot = New Pivot(90, NameOf(PivotTableLayoutBlankRows))
        Public Shared ReadOnly Property PivotTableLayoutGrandTotals As Pivot = New Pivot(91, NameOf(PivotTableLayoutGrandTotals))
        Public Shared ReadOnly Property PivotTableLayoutReportLayout As Pivot = New Pivot(92, NameOf(PivotTableLayoutReportLayout))
        Public Shared ReadOnly Property PivotTableLayoutShowInCompactForm As Pivot = New Pivot(93, NameOf(PivotTableLayoutShowInCompactForm))
        Public Shared ReadOnly Property PivotTableLayoutShowInOutlineForm As Pivot = New Pivot(94, NameOf(PivotTableLayoutShowInOutlineForm))
        Public Shared ReadOnly Property PivotTableLayoutShowInTabularForm As Pivot = New Pivot(95, NameOf(PivotTableLayoutShowInTabularForm))
        Public Shared ReadOnly Property PivotTableLayoutSubtotals As Pivot = New Pivot(96, NameOf(PivotTableLayoutSubtotals))
        Public Shared ReadOnly Property PivotTableListFormulas As Pivot = New Pivot(97, NameOf(PivotTableListFormulas))
        Public Shared ReadOnly Property PivotTableMove As Pivot = New Pivot(98, NameOf(PivotTableMove))
        Public Shared ReadOnly Property PivotTableNewStyle As Pivot = New Pivot(99, NameOf(PivotTableNewStyle))
        Public Shared ReadOnly Property PivotTableOlapConvertToFormulas As Pivot = New Pivot(100, NameOf(PivotTableOlapConvertToFormulas))
        Public Shared ReadOnly Property PivotTableOlapPropertyFields As Pivot = New Pivot(101, NameOf(PivotTableOlapPropertyFields))
        Public Shared ReadOnly Property PivotTableOlapTools As Pivot = New Pivot(102, NameOf(PivotTableOlapTools))
        Public Shared ReadOnly Property PivotTableOptions As Pivot = New Pivot(103, NameOf(PivotTableOptions))
        Public Shared ReadOnly Property PivotTableOptionsMenu As Pivot = New Pivot(104, NameOf(PivotTableOptionsMenu))
        Public Shared ReadOnly Property PivotTableReport As Pivot = New Pivot(105, NameOf(PivotTableReport))
        Public Shared ReadOnly Property PivotTableSelectData As Pivot = New Pivot(106, NameOf(PivotTableSelectData))
        Public Shared ReadOnly Property PivotTableSelectFlyout As Pivot = New Pivot(107, NameOf(PivotTableSelectFlyout))
        Public Shared ReadOnly Property PivotTableSelectLabel As Pivot = New Pivot(108, NameOf(PivotTableSelectLabel))
        Public Shared ReadOnly Property PivotTableSelectLabelAndData As Pivot = New Pivot(109, NameOf(PivotTableSelectLabelAndData))
        Public Shared ReadOnly Property PivotTableShowPages As Pivot = New Pivot(110, NameOf(PivotTableShowPages))
        Public Shared ReadOnly Property PivotTableStylesGallery As Pivot = New Pivot(111, NameOf(PivotTableStylesGallery))
        Public Shared ReadOnly Property PivotTableSubtotalsDoNotShow As Pivot = New Pivot(112, NameOf(PivotTableSubtotalsDoNotShow))
        Public Shared ReadOnly Property PivotTableSubtotalsOnBottom As Pivot = New Pivot(113, NameOf(PivotTableSubtotalsOnBottom))
        Public Shared ReadOnly Property PivotTableSubtotalsOnTop As Pivot = New Pivot(114, NameOf(PivotTableSubtotalsOnTop))
        Public Shared ReadOnly Property PivotUngroupItems As Pivot = New Pivot(115, NameOf(PivotUngroupItems))
        Public Shared ReadOnly Property TableExportTableToVisioPivotDiagram As Pivot = New Pivot(116, NameOf(TableExportTableToVisioPivotDiagram))
        Public Shared ReadOnly Property TableSummarizeWithPivot As Pivot = New Pivot(117, NameOf(TableSummarizeWithPivot))
        Public Shared ReadOnly Property ViewsPivotChartView As Pivot = New Pivot(118, NameOf(ViewsPivotChartView))
        Public Shared ReadOnly Property ViewsPivotTableView As Pivot = New Pivot(119, NameOf(ViewsPivotTableView))

        Public Overrides Function Clone() As Object
            Return New Pivot(value, name)
        End Function

    End Class

End Namespace