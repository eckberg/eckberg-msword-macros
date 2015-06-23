Attribute VB_Name = "Eckberg MS Word Placeholder Macro"

Sub addEmptyPlaceholder()
    
    Selection.TypeText Text:="[]"
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.InsertSymbol Font:="Verdana", CharacterNumber:=8226, Unicode:= _
        True
    Selection.MoveLeft Unit:=wdCharacter, Count:=3
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Options.DefaultHighlightColorIndex = wdYellow
    Selection.Range.HighlightColorIndex = wdYellow
    Selection.MoveRight Unit:=wdCharacter, Count:=1

End Sub
