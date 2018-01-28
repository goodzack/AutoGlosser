Sub StraightDupeZG1()

' StraightDupe
' I THINK THIS GRABS JUST THE LINE
Selection.StartOf Unit:=wdParagraph
Selection.MoveEnd Unit:=wdParagraph

With Selection.Find
  Selection.Copy
  Selection.Paste
  Selection.Paste
  Selection.MoveUp Unit:=wdParagraph, Count:=2, Extend:=wdExtend
  Selection.StartOf Unit:=wdParagraph
  Selection.MoveEnd Unit:=wdParagraph
  Selection.Find.ClearFormatting
  Selection.Find.Replacement.Style = ActiveDocument.Styles("Orthography")
End With

' TURNS HYPHENS TO EQUALS- MIGHT WANT TO TURN OFF SOMETIMES

With Selection.Find
    .ClearFormatting
    .Font.Name = “Straight”
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
    .Text = "ß"
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
    .Style = "Orthography"
    .Text = "c"
    .Replacement.Text = "ts"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "ç"
    .Style = "Orthography"
    .Replacement.Text = "ts’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "∂"
    .Replacement.Text = "ch’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "txø"
    .Replacement.Text = "t-hw"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Font.Name = “Straight”
    .Text = "tl"
    .Replacement.Text = "t-l"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "txø"
    .Replacement.Text = "t-hw"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "t÷"
    .Replacement.Text = "t-’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "p÷"
    .Replacement.Text = "p-’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "q÷"
    .Replacement.Text = "q-’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "c÷"
    .Replacement.Text = "c-’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "kø-÷"
    .Replacement.Text = "kw-’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "q÷"
    .Replacement.Text = "q-’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = ChrW(339) & ChrW(248)
    .Replacement.Text = "qw’"

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
    .Replacement.Text = "tth’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = ChrW(181)
    .Replacement.Text = "m’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Style = "Orthography"
    .Text = "©"
    .Replacement.Text = "tth"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Style = "Orthography"
    .Text = "∑"
    .Replacement.Text = "w’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Style = "Orthography"
    .Text = ChrW(931)
    .Replacement.Text = "w’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Style = "Orthography"
    .Text = "ø"
    .Replacement.Text = "xw"

    .Execute Replace:=wdReplaceAll
End With


With Selection.Find
    .ClearFormatting
    .Text = "÷"
    .Replacement.Text = "’"

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
    .Text = "∆"
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
    .Text = "xø"
    .Replacement.Text = "hw"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i"
    .Replacement.Text = "i"

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
    .Text = "k"
    .Replacement.Text = "k"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "kø"
    .Replacement.Text = "kw"

    .Execute Replace:=wdReplaceAll
End With

'ch' through kw

With Selection.Find
    .ClearFormatting
    .Text = ChrW(730) & ChrW(248)
    .Replacement.Text = "kw’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "¬"
    .Replacement.Text = "l’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "®"
    .Replacement.Text = "lh"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = ChrW(956)
    .Replacement.Text = "m’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "∫"
    .Replacement.Text = "n’"

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
    .Text = "π"
    .Replacement.Text = "p’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "œ"
    .Replacement.Text = "q’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "qø"
    .Replacement.Text = "qw"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "ß"
    .Replacement.Text = "sh"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "†"
    .Replacement.Text = "t’"

    .Execute Replace:=wdReplaceAll
End With


With Selection.Find
    .ClearFormatting
    .Text = "ƒ"
    .Replacement.Text = "th"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "√"
    .Replacement.Text = "tl’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "tƒ"
    .Style = "Orthography"
    .Replacement.Text = "ts’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = ChrW(32) & ChrW(775)
    .Style = "Orthography"
    .Replacement.Text = "tth’"

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
    .Replacement.Text = "w’"

    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "≈"
    .Style = "Orthography"
    .Replacement.Text = "x"
    .Execute Replace:=wdReplaceAll

End With

With Selection.Find
    .ClearFormatting
    .Text = "¥"
    .Style = "Orthography"
    .Replacement.Text = "y’"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "©"
    .Style = "Orthography"
    .Replacement.Text = "tth"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = ChrW(730) & ChrW(248)
    .Style = "Orthography"
    .Replacement.Text = “kw’”
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "e" & ChrW(181) & "i"
    .Replacement.Text = "e’mi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "e∫i"
    .Replacement.Text = "e’ni"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "e¬i"
    .Replacement.Text = "e’li"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = “e” & ChrW(8776) & “i”
    .Replacement.Text = "e’wi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "e¥i"
    .Replacement.Text = "e’yi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i" & ChrW(181) & "i"
    .Replacement.Text = "i’mi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i∫i"
    .Replacement.Text = "i’ni"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i¬i"
    .Replacement.Text = "i’li"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Style = "Orthography"
    .Text = "i∑i"
    .Replacement.Text = "i’wi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i¥i"
    .Replacement.Text = "i’yi"
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
    .Replacement.Text = "tth’"
    .Execute Replace:=wdReplaceAll
End With

' distributing glottal for m

With Selection.Find
    .ClearFormatting
    .Text = "em’i"
    .Replacement.Text = "e’mi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i’me"
    .Replacement.Text = "e’mi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "am’i"
    .Replacement.Text = "a’mi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i’ma"
    .Replacement.Text = "im’a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "em’u"
    .Replacement.Text = "e’mu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u’me"
    .Replacement.Text = "um’e"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u’ma"
    .Replacement.Text = "um’a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "am’u"
    .Replacement.Text = "a’mu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "im’u"
    .Replacement.Text = "i’mu"
    .Execute Replace:=wdReplaceAll
End With

' distributing glottal for n

With Selection.Find
    .ClearFormatting
    .Text = "en’i"
    .Replacement.Text = "e’ni"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i’ne"
    .Replacement.Text = "e’ni"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "an’i"
    .Replacement.Text = "a’ni"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i’na"
    .Replacement.Text = "in’a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "en’u"
    .Replacement.Text = "e’nu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u’ne"
    .Replacement.Text = "un’e"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u’na"
    .Replacement.Text = "un’a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "an’u"
    .Replacement.Text = "a’nu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "in’u"
    .Replacement.Text = "i’nu"
    .Execute Replace:=wdReplaceAll
End With

' distributing glottal for l

With Selection.Find
    .ClearFormatting
    .Text = "el’i"
    .Replacement.Text = "e’li"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i’le"
    .Replacement.Text = "e’li"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "al’i"
    .Replacement.Text = "a’li"
    .Execute Replace:=wdReplaceAll

End With

With Selection.Find
    .ClearFormatting
    .Text = "i’la"
    .Replacement.Text = "il’a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "el’u"
    .Replacement.Text = "e’lu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u’le"
    .Replacement.Text = "ul’e"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u’la"
    .Replacement.Text = "ul’a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "al’u"
    .Replacement.Text = "a’lu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "il’u"
    .Replacement.Text = "i’lu"
    .Execute Replace:=wdReplaceAll
End With

' distributing glottal for y

With Selection.Find
    .ClearFormatting
    .Text = "ey’i"
    .Replacement.Text = "e’yi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i’ye"
    .Replacement.Text = "e’yi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "ay’i"
    .Replacement.Text = "a’yi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i’ya"
    .Replacement.Text = "iy’a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "ey’u"
    .Replacement.Text = "e’yu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u’ye"
    .Replacement.Text = "uy’e"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u’ya"
    .Replacement.Text = "uy’a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "ay’u"
    .Replacement.Text = "a’yu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "iy’u"
    .Replacement.Text = "i’yu"
    .Execute Replace:=wdReplaceAll
End With

' distributing glottal for w

With Selection.Find
    .ClearFormatting
    .Text = "ew’i"
    .Replacement.Text = "e’wi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i’we"
    .Replacement.Text = "e’wi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "aw’i"
    .Replacement.Text = "a’wi"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "i’wa"
    .Replacement.Text = "iw’a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "ew’u"
    .Replacement.Text = "e’wu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u’we"
    .Replacement.Text = "uw’e"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "u’wa"
    .Replacement.Text = "uw’a"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "aw’u"
    .Replacement.Text = "a’wu"
    .Execute Replace:=wdReplaceAll
End With

With Selection.Find
    .ClearFormatting
    .Text = "iw’u"
    .Replacement.Text = "i’wu"
    .Execute Replace:=wdReplaceAll
End With

' for some reason, the macro turns "True" to "0". This should fix it but will also screw with anytime you have the string "True".

With Selection.Find
    .ClearFormatting
    .Text = “True“
    .Replacement.Text = "0"
    .Execute Replace:=wdReplaceAll
End With

End Sub
