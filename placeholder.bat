Attribute VB_Name = "WeComply MS Word Placeholder Macro"

Sub WeComplyMsWordPlaceholder()
    
    Selection.TypeText Text:="["
    Selection.InsertSymbol Font:="Verdana", CharacterNumber:=8226, Unicode:=True
    Selection.TypeText Text:="]"
    Selection.MoveLeft Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Options.DefaultHighlightColorIndex = wdYellow
    Selection.Range.HighlightColorIndex = wdYellow

End Sub
