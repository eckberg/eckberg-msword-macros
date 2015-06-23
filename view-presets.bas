Attribute VB_Name = "Eckberg MS Word View Preset Macros"

Public Sub viewPresetFullWidth()
    
    ActiveDocument.ActiveWindow.WindowState = wdWindowStateMaximize
    ActiveDocument.ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitBestFit
    
    ActiveDocument.ActiveWindow.DocumentMap = True
    ActiveDocument.Application.CommandBars("Navigation").Position = msoBarRight
    ActiveDocument.Application.CommandBars("Navigation").Width = ActiveDocument.Application.Width / 5
    
    ActiveDocument.Application.TaskPanes(wdTaskPaneFormatting).Visible = True
    ActiveDocument.Application.CommandBars("Styles").Position = msoBarLeft
    ActiveDocument.Application.CommandBars("Styles").Width = ActiveDocument.Application.Width / 5
    
End Sub

Public Sub viewPresetFloatWidth()
    
    ActiveDocument.ActiveWindow.WindowState = wdWindowStateNormal
    ActiveDocument.ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitBestFit
    ActiveDocument.Application.Width = Application.UsableWidth / 1.5
    ActiveDocument.Application.Height = Application.UsableHeight / 1.15
    
    ActiveDocument.ActiveWindow.DocumentMap = False
        
    ActiveDocument.Application.TaskPanes(wdTaskPaneFormatting).Visible = False
    
End Sub
