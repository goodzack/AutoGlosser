Attribute VB_Name = "RawPower"



Sub TabAndDupeZG4()

'If Selection.Type = wdSelectionIP Then
'    MsgBox ("Please select a line of data before running this macro.")
Selection.StartOf Unit:=wdParagraph
Selection.MoveEnd Unit:=wdParagraph
'Else

' Clear tab stops and reset default to .5"
    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = InchesToPoints(0.5)

'Removes Double Spaces
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

'Converts Spaces to Tabs
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " "
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Collapse Direction:=wdCollapseStart
    Selection.TypeText "#"
    Selection.TypeText vbTab
    Selection.HomeKey Unit:=wdLine
    Selection.MoveEnd Unit:=wdParagraph

' Duplicates Line
Selection.Style = ActiveDocument.Styles("Gloss")
Selection.Font.Color = wdColorTeal
Selection.Copy
Selection.Paste
Selection.Paste
Selection.MoveUp Unit:=wdParagraph, Count:=2, Extend:=wdExtend
Selection.StartOf Unit:=wdParagraph
Selection.MoveEnd Unit:=wdParagraph
Selection.Style = ActiveDocument.Styles("Phonetics-Hyphen")
Selection.Font.Color = wdAutomatic



'Insert tab at end of line
    Selection.Find.ClearFormatting
    Selection.Find.Font.Color = wdColorTeal

    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^t#^p"
        .Forward = False
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
'End If

End Sub

Sub StraightDupeZG1()

' StraightDupe

Selection.StartOf Unit:=wdParagraph
Selection.MoveEnd Unit:=wdParagraph

With Selection.Find
  Selection.Style = ActiveDocument.Styles("Phonetics")
    Selection.Find.Execute Replace:=wdReplaceAll
  Selection.Copy
  Selection.Paste
  Selection.Paste
  Selection.MoveUp Unit:=wdParagraph, Count:=2, Extend:=wdExtend
  Selection.StartOf Unit:=wdParagraph
  Selection.MoveEnd Unit:=wdParagraph
  Selection.ClearFormatting
  Selection.Style = ActiveDocument.Styles("Orthography")
    Selection.Find.Execute Replace:=wdReplaceAll
End With

' TURNS HYPHENS TO EQUALS- MIGHT WANT TO TURN OFF SOMETIMES

With Selection.Find
    .ClearFormatting
    .Font.Name = �Straight�
    .Text = ChrW(45)
    .Replacement.Text = ChrW(61)

    .Execute Replace:=wdReplaceAll
End With

'THIS DELETES ACCENTS - MIGHT WANT TO TURN OFF SOMETIMES


With Selection.Find
    .ClearFormatting
    .Text = ChrW(233)
    .Replacement.Text = "e"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = ChrW(225)
    .Replacement.Text = "a"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = ChrW(237)
    .Replacement.Text = "i"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "sh"
    .Replacement.Text = "s-h"

    .Execute Replace:=wdReplaceAll
End With


With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Replacement.Text = "sh"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Style = "Orthography"
    .Text = "ts"
    .Replacement.Text = "t-s"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "c"
    .Replacement.Text = "ts"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Replacement.Text = "ts�"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Replacement.Text = "ch�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "tx�"
    .Replacement.Text = "t-hw"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "��"
    .Replacement.Text = "kw�"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Replacement.Text = "q�"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Font.Name = �Straight�
    .Text = "tl"
    .Replacement.Text = "t-l"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "tx�"
    .Replacement.Text = "t-hw"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "t�"
    .Replacement.Text = "t-�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "p�"
    .Replacement.Text = "p-�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "q�"
    .Replacement.Text = "q-�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "c�"
    .Replacement.Text = "c-�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "k�-�"
    .Replacement.Text = "kw-�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "q�"
    .Replacement.Text = "q-�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = ChrW(339) & ChrW(248)
    .Replacement.Text = "qw�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = ChrW(8776) & ChrW(248)
    .Style = "Orthography"
    .Replacement.Text = "xw"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Font.Name = "Times"
    .Text = ChrW(730)
    .Replacement.Text = "tth�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = ChrW(181)
    .Replacement.Text = "m�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Style = "Orthography"
    .Text = "�"
    .Replacement.Text = "tth"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Style = "Orthography"
    .Text = "�"
    .Replacement.Text = "w�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Style = "Orthography"
    .Text = ChrW(931)
    .Replacement.Text = "w�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Style = "Orthography"
    .Text = "�"
    .Replacement.Text = "w"
    .Execute Replace:=wdReplaceAll
End With


With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Replacement.Text = "�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Style = "Orthography"
    .Text = "a:"
    .Replacement.Text = "aa"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Replacement.Text = "ch"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Style = "Orthography"
    .Text = "e:"
    .Replacement.Text = "ee"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "x�"
    .Replacement.Text = "hw"

    .Execute Replace:=wdReplaceAll
End With


With Selection.Find
    .ClearFormatting
    .Style = "Orthography"
    .Text = "i:"
    .Replacement.Text = "ii"

    .Execute Replace:=wdReplaceAll
End With


With Selection.Find
    .ClearFormatting
    .Text = "k�"
    .Replacement.Text = "kw"

    .Execute Replace:=wdReplaceAll
End With

'ch' through kw

With Selection.Find
    .ClearFormatting
    .Text = ChrW(730) & ChrW(248)
    .Replacement.Text = "kw�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Replacement.Text = "l�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Replacement.Text = "lh"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = ChrW(956)
    .Replacement.Text = "m�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Replacement.Text = "n�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u:"
    .Replacement.Text = "oo"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u"
    .Style = "Orthography"
    .Replacement.Text = "ou"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Replacement.Text = "p�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Replacement.Text = "q�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "q�"
    .Replacement.Text = "qw"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Replacement.Text = "sh"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Replacement.Text = "t�"

    .Execute Replace:=wdReplaceAll
End With


With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Replacement.Text = "th"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Replacement.Text = "tl�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "t�"
    .Style = "Orthography"
    .Replacement.Text = "ts�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = ChrW(32) & ChrW(775)
    .Style = "Orthography"
    .Replacement.Text = "tth�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Style = "Orthography"
    .Text = ChrW(59)
    .Replacement.Text = "u"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = ChrW(8721)
    .MatchCase = True
    .Style = "Orthography"
    .Replacement.Text = "w�"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Style = "Orthography"
    .Replacement.Text = "x"
    .Execute Replace:=wdReplaceAll

End With

With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Style = "Orthography"
    .Replacement.Text = "y�"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "�"
    .Replacement.Text = "tth"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = ChrW(730) & ChrW(248)
    .Replacement.Text = �kw��
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "e" & ChrW(181) & "i"
    .Replacement.Text = "e�mi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "e�i"
    .Replacement.Text = "e�ni"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "e�i"
    .Replacement.Text = "e�li"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = �e� & ChrW(8776) & �i�
    .Replacement.Text = "e�wi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "e�i"
    .Replacement.Text = "e�yi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i" & ChrW(181) & "i"
    .Replacement.Text = "i�mi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i�i"
    .Replacement.Text = "i�ni"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i�i"
    .Replacement.Text = "i�li"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Style = "Orthography"
    .Text = "i�i"
    .Replacement.Text = "i�wi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i�i"
    .Replacement.Text = "i�yi"
    .Execute Replace:=wdReplaceAll
End With

' the following change the English ts/tl-digraphs etc. with a hyphen back to how they should be. FIX LATER: target fonts that AREN'T straight, instead of targeting Times.

With Selection.Find
    .ClearFormatting
    .Font.Name = "Times"
    .Text = "t-s"
    .Replacement.Text = "ts"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Font.Name = "Times"
    .Text = "t-l"
    .Replacement.Text = "tl"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Font.Name = "Times"
    .Text = "s" & "-" & "h"
    .Replacement.Text = "sh"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = ChrW(729)
    .Replacement.Text = "tth�"
    .Execute Replace:=wdReplaceAll
End With

' distributing glottal for m

With Selection.Find
    .ClearFormatting
    .Text = "em�i"
    .Replacement.Text = "e�mi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i�me"
    .Replacement.Text = "e�mi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "am�i"
    .Replacement.Text = "a�mi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i�ma"
    .Replacement.Text = "im�a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "em�u"
    .Replacement.Text = "e�mu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u�me"
    .Replacement.Text = "um�e"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u�ma"
    .Replacement.Text = "um�a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "am�u"
    .Replacement.Text = "a�mu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "im�u"
    .Replacement.Text = "i�mu"
    .Execute Replace:=wdReplaceAll
End With

' distributing glottal for n

With Selection.Find
    .ClearFormatting
    .Text = "en�i"
    .Replacement.Text = "e�ni"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i�ne"
    .Replacement.Text = "e�ni"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "an�i"
    .Replacement.Text = "a�ni"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i�na"
    .Replacement.Text = "in�a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "en�u"
    .Replacement.Text = "e�nu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u�ne"
    .Replacement.Text = "un�e"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u�na"
    .Replacement.Text = "un�a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "an�u"
    .Replacement.Text = "a�nu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "in�u"
    .Replacement.Text = "i�nu"
    .Execute Replace:=wdReplaceAll
End With

' distributing glottal for l

With Selection.Find
    .ClearFormatting
    .Text = "el�i"
    .Replacement.Text = "e�li"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i�le"
    .Replacement.Text = "e�li"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "al�i"
    .Replacement.Text = "a�li"
    .Execute Replace:=wdReplaceAll

End With

With Selection.Find
    .ClearFormatting
    .Text = "i�la"
    .Replacement.Text = "il�a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "el�u"
    .Replacement.Text = "e�lu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u�le"
    .Replacement.Text = "ul�e"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u�la"
    .Replacement.Text = "ul�a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "al�u"
    .Replacement.Text = "a�lu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "il�u"
    .Replacement.Text = "i�lu"
    .Execute Replace:=wdReplaceAll
End With

' distributing glottal for y

With Selection.Find
    .ClearFormatting
    .Text = "ey�i"
    .Replacement.Text = "e�yi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i�ye"
    .Replacement.Text = "e�yi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "ay�i"
    .Replacement.Text = "a�yi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i�ya"
    .Replacement.Text = "iy�a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "ey�u"
    .Replacement.Text = "e�yu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u�ye"
    .Replacement.Text = "uy�e"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u�ya"
    .Replacement.Text = "uy�a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "ay�u"
    .Replacement.Text = "a�yu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "iy�u"
    .Replacement.Text = "i�yu"
    .Execute Replace:=wdReplaceAll
End With

' distributing glottal for w

With Selection.Find
    .ClearFormatting
    .Text = "ew�i"
    .Replacement.Text = "e�wi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i�we"
    .Replacement.Text = "e�wi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "aw�i"
    .Replacement.Text = "a�wi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i�wa"
    .Replacement.Text = "iw�a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "ew�u"
    .Replacement.Text = "e�wu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u�we"
    .Replacement.Text = "uw�e"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u�wa"
    .Replacement.Text = "uw�a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "aw�u"
    .Replacement.Text = "a�wu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "iw�u"
    .Replacement.Text = "i�wu"
    .Execute Replace:=wdReplaceAll
End With

' for some reason, the macro turns "True" to "0". This should fix it but will also screw with anytime you have the string "True".

With Selection.Find
    .ClearFormatting
    .Text = �True�
    .Replacement.Text = "0"
    .Execute Replace:=wdReplaceAll
End With

End Sub

Sub PrepFourLineZG4()

'This selects the whole paragraph the cursor in.
Selection.StartOf Unit:=wdParagraph
Selection.MoveEnd Unit:=wdParagraph

' Clear tab stops and reset default to .5"
    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = InchesToPoints(0.5)

'Removes Double Spaces
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

' Duplicates Line
Selection.Font.Color = wdColorGray50
Selection.Copy
Selection.Paste
Selection.Paste
Selection.Collapse Direction:=wdCollapseStart
Selection.MoveUp Unit:=wdParagraph, Count:=2
Selection.StartOf Unit:=wdParagraph
Selection.MoveEnd Unit:=wdParagraph
Selection.Style = "Phonetics"

''Highlight next suspected line for running the macro on


'' Change copied lines back to auto (black)
'    Selection.Find.ClearFormatting
'    Selection.Find.Font.Color = wdColorDarkYellow
'    Selection.Find.Style = "Orthography"
'    Selection.Find.Replacement.ClearFormatting
'    Selection.Find.Replacement.Font.Color = wdColorAutomatic
'    With Selection.Find
'        .Text = ""
'        .Replacement.Text = ""
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = True
'        .MatchCase = False
'        .MatchWholeWord = False
'        .MatchWildcards = False
'        .MatchSoundsLike = False
'        .MatchAllWordForms = False
'    End With
'    Selection.Find.Execute Replace:=wdReplaceAll

End Sub



Sub RawPower()

Call StraightDupeZG1
 
 Selection.Collapse Direction:=wdCollapseEnd
 Selection.MoveDown Unit:=wdParagraph, Count:=3


End Sub
