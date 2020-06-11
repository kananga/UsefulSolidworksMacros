Option Explicit
'   Required reference: Microsoft Scripting Runtime

Const scalePercent As Double = 10

Sub ScaleUp()
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    Dim swDoc As SldWorks.ModelDoc2
    Set swDoc = swApp.ActiveDoc
    
    If swDoc Is Nothing Then
        Exit Sub
    ElseIf swDoc.GetType <> swDocumentTypes_e.swDocDRAWING Then
        Exit Sub
    End If
    
    updateScale swDoc, 1 + scalePercent / 100
    swDoc.Extension.Rebuild swRebuildOptions_e.swCurrentSheetDisp
End Sub

Sub ScaleDown()
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    Dim swDoc As SldWorks.ModelDoc2
    Set swDoc = swApp.ActiveDoc
    
    If swDoc Is Nothing Then
        Exit Sub
    ElseIf swDoc.GetType <> swDocumentTypes_e.swDocDRAWING Then
        Exit Sub
    End If
    
    updateScale swDoc, 1 / (1 + scalePercent / 100)
    swDoc.Extension.Rebuild swRebuildOptions_e.swCurrentSheetDisp
End Sub

Private Sub updateScale(ByRef swDoc As SldWorks.ModelDoc2, ByRef scaleDecimalMultiplier As Double)
    Dim swDrawing As SldWorks.DrawingDoc
    Set swDrawing = swDoc
    
    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDrawing.GetCurrentSheet
    
    Dim viewsToScale As Dictionary
    Set viewsToScale = getSelectedViews(swSheet, swDoc)
    
    If viewsToScale.Count = 0 Then
        Set viewsToScale = getCurrentSheetViews(swSheet)
    End If
    
    Dim viewName As Variant
    Dim swView As SldWorks.View
    
    For Each viewName In viewsToScale
        Set swView = viewsToScale(CStr(viewName))
        swView.ScaleDecimal = swView.ScaleDecimal * scaleDecimalMultiplier
    Next viewName
End Sub

Private Function getSelectedViews(ByRef swSheet As SldWorks.Sheet, ByRef swDoc As SldWorks.ModelDoc2) As Dictionary
    Set getSelectedViews = New Dictionary

    Dim swSelectionMgr As SldWorks.SelectionMgr
    Set swSelectionMgr = swDoc.SelectionManager
    
    If swSelectionMgr.GetSelectedObjectCount2(-1) = 0 Then
        Exit Function
    End If
    
    Dim swView As SldWorks.View
    Dim i As Integer
    
    For i = 1 To swSelectionMgr.GetSelectedObjectCount2(-1)
        If swSelectionMgr.GetSelectedObjectType3(i, -1) = swSelectType_e.swSelDRAWINGVIEWS Then
            Set swView = swSelectionMgr.GetSelectedObject6(i, -1)
            If swView.Sheet.GetName = swSheet.GetName And Not getSelectedViews.Exists(swView.Name) Then
                getSelectedViews.Add swView.Name, swView
            End If
        End If
    Next i
End Function

Private Function getCurrentSheetViews(ByRef swSheet As SldWorks.Sheet) As Dictionary
    Set getCurrentSheetViews = New Dictionary
    
    Dim views As Variant
    views = swSheet.GetViews

    Dim viewVariant As Variant
    Dim swView As SldWorks.View

    For Each viewVariant In views
        Set swView = viewVariant
        If swView.SuppressState = 0 Then
            If Not getCurrentSheetViews.Exists(swView.Name) Then
                getCurrentSheetViews.Add swView.Name, swView
            End If
        End If
    Next viewVariant
End Function
