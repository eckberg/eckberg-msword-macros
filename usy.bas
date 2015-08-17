Attribute VB_Name = "USY MS Word Macro"

Sub usy()

    Dim range As range
    
    Set range = ActiveDocument.Content
    With range.Find
        .ClearFormatting
        .MatchWholeWord = True
        .MatchCase = False
        .Text = "skall"
    End With
    Do While range.Find.Execute = True
        ActiveDocument.Comments.Add range, "Använd ska istället för skall."
    Loop
    
    Set range = ActiveDocument.Content
    With range.Find
        .ClearFormatting
        .MatchWholeWord = False
        .MatchCase = False
        .Text = "orgnr."
    End With
    Do While range.Find.Execute = True
        ActiveDocument.Comments.Add range, "Använd förkortningen org.nr."
    Loop
        
End Sub
