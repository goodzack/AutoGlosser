Attribute VB_Name = "AutoGlosserModules"

Sub TabAndDupeZP3()

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
Selection.Copy
Selection.Paste
Selection.Paste
Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
Selection.StartOf Unit:=wdParagraph
Selection.MoveEnd Unit:=wdParagraph
Selection.Font.Color = wdColorTeal

'Insert tab at end of line
    Selection.Find.ClearFormatting
    Selection.Find.Font.Color = wdColorTeal
    Selection.Find.Replacement.ClearFormatting
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

''Highlight next suspected line for running the macro on
Selection.Collapse Direction:=wdCollapseEnd
Selection.Find.ClearFormatting
    With Selection.Find
        .Font.Name = "Straight"
        .Font.Color = wdColorAutomatic
        .Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    If Selection.Find.Found = True Then
    Selection.StartOf Unit:=wdParagraph
    Selection.MoveEnd Unit:=wdParagraph
End If
End Sub

Sub DangerRobotZP3()
Attribute DangerRobotZP3.VB_ProcData.VB_Invoke_Func = "Normal.AutoGlosserModules.DangerRobotZP3"

Dim intMsgBoxResult As Integer
intMsgBoxResult = MsgBox("This macro is quite drastic and is only intended for files that have been carefully prepared. Are you sure you want to proceed?", vbYesNo + _
   vbQuestion, "Proceed With Caution")
   If intMsgBoxResult = vbYes Then
    
intMsgBoxResult = MsgBox("This action is essentially irreversible. Have you made a backup copy of this file? If not, please select 'No' and then use File>Save As to make a copy. If you already have a backup copy and are ready to proceed, select 'Yes'.", vbYesNo + _
   vbQuestion, "Did You Keep A Copy?")
   If intMsgBoxResult = vbYes Then
   
Selection.Find.ClearFormatting
    With Selection.Find
        .Font.Name = "Straight"
        .Font.Color = wdColorAutomatic
        .Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Do Until Selection.Find.Found = False
    If Selection.Find.Found = True Then
    Call TabAndDupeZP3
    End If
    Loop
Call GramMorphGlosserZP3
End If
End If
End Sub


Sub DoFindReplaceTeal(FindText As String, ReplaceText As String)

With Selection.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = FindText
    .Replacement.Text = ReplaceText
    
    Selection.Find.Font.Color = wdColorTeal

    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    Do While .Execute
        'Keep going until nothing found
        .Execute Replace:=wdReplaceAll
    Loop
    'Free up some memory
    ActiveDocument.UndoClear
End With

End Sub
Sub DoFindReplaceAsString(FindText As String, ReplaceText As String)
Attribute DoFindReplaceAsString.VB_Description = "Macro created 5/13/09 by Zoey Peterson"
Attribute DoFindReplaceAsString.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.DoFindReplaceAsString"
'
' DoFindReplaceAsString Macro
' Macro created 5/13/09 by Zoey Peterson
'Sub DoFindReplace

With Selection.Find
    .ClearFormatting
    .Replacement.ClearFormatting

    .Text = FindText
    .Replacement.Text = ReplaceText

    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    Do While .Execute
        'Keep going until nothing found
        .Execute Replace:=wdReplaceAll
    Loop
    'Free up some memory
    ActiveDocument.UndoClear
End With


End Sub

Sub PrepFourLineZP3()

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
Selection.Copy
Selection.Paste
Selection.Paste
Selection.Collapse Direction:=wdCollapseStart
Selection.MoveUp Unit:=wdParagraph, Count:=2
Selection.StartOf Unit:=wdParagraph
Selection.MoveEnd Unit:=wdParagraph
Selection.Font.Color = wdColorGray50

''Highlight next suspected line for running the macro on
Selection.Collapse Direction:=wdCollapseEnd
Selection.MoveDown Unit:=wdParagraph, Count:=3
Selection.Find.ClearFormatting
    With Selection.Find
        .Font.Name = "Straight"
        .Font.Color = wdColorAutomatic
        .Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    If Selection.Find.Found = True Then
    Selection.StartOf Unit:=wdParagraph
    Selection.MoveEnd Unit:=wdParagraph
End If

'' Change copied lines back to auto (black)
'    Selection.Find.ClearFormatting
'    Selection.Find.Font.Color = wdColorDarkYellow
'    Selection.Find.Font.Name = "Straight"
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
Sub GramMorphGlosserZP3()

' GramMorphGlosser Macro
' Macro created 8/25/09 by Zoey Peterson
'
'MsgBox ("This might take a minute. I'll let you know when I'm done. Please click OK to begin.")

'Segment the hyphens
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
Selection.Find.Font.Color = wdColorTeal
    With Selection.Find
        .Text = "-"
        .Replacement.Text = "^t-^t"
        .Forward = False
        .Wrap = wdFindContinue
        .Format = True
     End With
    Selection.Find.Execute Replace:=wdReplaceAll

' Remove Punctuation
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
Selection.Find.Font.Color = wdColorTeal
    With Selection.Find
        .Text = "."
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
     End With
    Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
Selection.Find.Font.Color = wdColorTeal
    With Selection.Find
        .Text = ","
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
     End With
    Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
Selection.Find.Font.Color = wdColorTeal
    With Selection.Find
        .Text = "?"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = False
     End With
    Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
Selection.Find.Font.Color = wdColorTeal
    With Selection.Find
        .Text = "!"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = False
     End With
    Selection.Find.Execute Replace:=wdReplaceAll
   
' Replace Grammatical Morphemes
Call GramMorphEngine("^tels^t", "^tact^t")
Call GramMorphEngine("^t;¬s^t", "^tact^t")
Call GramMorphEngine("^tls^t", "^tact^t")
Call GramMorphEngine("^tni÷^t", "^taux^t")
Call GramMorphEngine("^t÷i^t", "^taux^t")
Call GramMorphEngine("^t÷i÷^t", "^tconj^t")
Call GramMorphEngine("^t;®c^t", "^tben^t")
Call GramMorphEngine("^tst;xø^t", "^tcs^t")
Call GramMorphEngine("^tstaµß^t", "^tcs:1obj^t")
Call GramMorphEngine("^t÷a÷l;^t", "^tcur^t")
Call GramMorphEngine("^t÷a¬;^t", "^tcur^t")
Call GramMorphEngine("^tn;s^t", "^tdir^t")
Call GramMorphEngine("^t˚ø;∫a^t", "^tdm^t")
Call GramMorphEngine("^t©e¥^t", "^tdm^t")
Call GramMorphEngine("^tt;÷i^t", "^tdm^t")
Call GramMorphEngine("^tt;÷i÷^t", "^tdm^t")
Call GramMorphEngine("^tkøƒe¥^t", "^tdm^t")
Call GramMorphEngine("^tt;∫a^t", "^tdm^t")
Call GramMorphEngine("^tƒ;^t", "^tdt^t")
Call GramMorphEngine("^t©;^t", "^tdt^t")
Call GramMorphEngine("^tkøƒ;^t", "^tdt^t")
Call GramMorphEngine("^tkø^t", "^tdt^t")
Call GramMorphEngine("^t˚ø^t", "^tdt^t")
Call GramMorphEngine("^t√^t", "^tdt^t")
Call GramMorphEngine("^t®;^t", "^tdt^t")
Call GramMorphEngine("^t©;∫^t", "^tdt:2pos^t")
Call GramMorphEngine("^tkøs^t", "^tdt:n^t")
Call GramMorphEngine("^tç;^t", "^tevid^t")
Call GramMorphEngine("^tce÷^t", "^tfut^t")
Call GramMorphEngine("^txø;^t", "^tinc^t")
Call GramMorphEngine("^tt;n^t", "^tins^t")
Call GramMorphEngine("^tn;xø^t", "^tlctr^t")
Call GramMorphEngine("^t÷;∑^t", "^tlnk^t")
Call GramMorphEngine("^tƒel;m^t", "^ttr:1pas^t")
Call GramMorphEngine("^t;ƒa:µ^t", "^ttr:2pas^t")
Call GramMorphEngine("^tn;^t", "^t1pos^t")
Call GramMorphEngine("^t;®^t", "^tpst^t")
Call GramMorphEngine("^t÷;®^t", "^tpst^t")
Call GramMorphEngine("^tƒ;t^t", "^trefl^t")
Call GramMorphEngine("^tƒat^t", "^trefl^t")
Call GramMorphEngine("^tme÷^t", "^trel^t")
Call GramMorphEngine("^ty;^t", "^tser^t")
Call GramMorphEngine("^tc;n^t", "^tsub^t")
Call GramMorphEngine("^ts;∑^t", "^tsub^t")
Call GramMorphEngine("^t∆^t", "^t2sub^t")
Call GramMorphEngine("^tce:p^t", "^t2pl.sub^t")
Call GramMorphEngine("^t;t^t", "^t1pl.ssub^t")
Call GramMorphEngine("^t;l;p^t", "^t2pl.ssub^t")
Call GramMorphEngine("^tt^t", "^ttr^t")
Call GramMorphEngine("^tye®^t", "^tseq^t")
Call GramMorphEngine("^txøi÷^t", "^tunexp^t")

'Replace Words
Call WordGlossaryEngine("^tpas^t", "^thit^t")
Call WordGlossaryEngine("^tm;stim;xø^t", "^tpeople^t")
Call WordGlossaryEngine("^txø;nit;µ^t", "^twhite.person^t")
Call WordGlossaryEngine("^tsw;œøa÷®^t", "^tmt.goat.blanket^t")
Call WordGlossaryEngine("^tt;m;xø^t", "^tearth^t")
Call WordGlossaryEngine("^tsqø;me¥,^t", "^tdog^t")
Call WordGlossaryEngine("^tπ;œ^t", "^twhite^t")
Call WordGlossaryEngine("^t÷;w;^t", "^tnot^t")
Call WordGlossaryEngine("^t÷e:®t;n^t", "^tthey^t")
Call WordGlossaryEngine("^tneµ^t", "^tgo^t")
Call WordGlossaryEngine("^twa¬a÷^t", "^tmaybe^t")
Call WordGlossaryEngine("^tstem^t", "^twhat^t")
Call WordGlossaryEngine("^tha÷^t", "^tif^t")
Call WordGlossaryEngine("^ttec;l^t", "^tarrive^t")
Call WordGlossaryEngine("^tlel;µ^t", "^thouse^t")
Call WordGlossaryEngine("^tƒi^t", "^tbig^t")
Call WordGlossaryEngine("^tl;≈øt;n^t", "^tblanket^t")
Call WordGlossaryEngine("^tqe÷is^t", "^tnow^t")
Call WordGlossaryEngine("^tsi÷eµ^t", "^trespected^t")
Call WordGlossaryEngine("^tse¥^t", "^twool^t")
Call WordGlossaryEngine("^tq;≈^t", "^tlots.of^t")
Call WordGlossaryEngine("^tsa:÷®^t", "^tour^t")
Call WordGlossaryEngine("^ttel;^t", "^tmoney^t")
Call WordGlossaryEngine("^ts÷;¬el;xø^t", "^telders^t")
Call WordGlossaryEngine("^t√;liµ^t", "^treally^t")
Call WordGlossaryEngine("^ts÷a≈øa÷^t", "^tbutter.clam^t")
Call WordGlossaryEngine("^ts√i÷√q;®^t", "^tchild^t")
Call WordGlossaryEngine("^tn;w;^t", "^tyou^t")
Call WordGlossaryEngine("^tn;ça÷^t", "^tone^t")
Call WordGlossaryEngine("^thiƒ^t", "^tlong.time^t")
Call WordGlossaryEngine("^tµi^t", "^tcome^t")
Call WordGlossaryEngine("^tyaƒ^t", "^talways^t")
Call WordGlossaryEngine("^tç;xøle÷^t", "^tsometimes^t")
Call WordGlossaryEngine("^txø;∫^t", "^tstill^t")
Call WordGlossaryEngine("^t÷aµ;t^t", "^tat.home^t")
Call WordGlossaryEngine("^tp;∫e¬;≈;˙^t", "^tKuper.Island^t")
Call WordGlossaryEngine("^tpest;n^t", "^tUnited.States^t")
Call WordGlossaryEngine("^t®qe¬ç^t", "^tmoon^t")
Call WordGlossaryEngine("^tsƒ;qi^t", "^tsockeye^t")
Call WordGlossaryEngine("^t˚øa¬;xø^t", "^tdog.salmon^t")
Call WordGlossaryEngine("^tsnet^t", "^tnight^t")
Call WordGlossaryEngine("^t÷;√q;l^t", "^tgo.outside^t")
Call WordGlossaryEngine("^tput^t", "^tboat^t")
Call WordGlossaryEngine("^t˚øin^t", "^thow.many^t")
Call WordGlossaryEngine("^tqa÷^t", "^twater^t")
Call WordGlossaryEngine("^tq;¬et^t", "^tagain^t")
Call WordGlossaryEngine("^txø;lm;xø^t", "^tnative^t")
Call WordGlossaryEngine("^tlisek^t", "^tsack^t")
Call WordGlossaryEngine("^tswe:m^t", "^thorse.clam^t")
Call WordGlossaryEngine("^ts√;la÷;m^t", "^tcockle^t")
Call WordGlossaryEngine("^t≈øi¬;µ^t", "^trope^t")
Call WordGlossaryEngine("^tœam^t", "^tkelp^t")
Call WordGlossaryEngine("^tœø;l^t", "^tcook^t")
Call WordGlossaryEngine("^tyays^t", "^twork^t")
Call WordGlossaryEngine("^tskøey;l^t", "^tday^t")
Call WordGlossaryEngine("^tƒaƒ;n^t", "^tmouth^t")
Call WordGlossaryEngine("^t†a˚ø^t", "^tgo.home^t")
Call WordGlossaryEngine("^tç;¥xø^t", "^tdry^t")
Call WordGlossaryEngine("^tsce:®t;n^t", "^tsalmon^t")
Call WordGlossaryEngine("^th;¥qø^t", "^tfire^t")
Call WordGlossaryEngine("^tspe:nxø^t", "^tcamas^t")
Call WordGlossaryEngine("^tsa:¥^t", "^tready^t")
Call WordGlossaryEngine("^t≈ƒ;m^t", "^tbox^t")
Call WordGlossaryEngine("^txøm;ƒkøi÷;m^t", "^tMusqueam^t")
Call WordGlossaryEngine("^tçtwa÷^t", "^tperhaps^t")
Call WordGlossaryEngine("^tßxø;µnikø^t", "^taunt/uncle^t")
Call WordGlossaryEngine("^tlel;µ^t", "^thouse^t")
Call WordGlossaryEngine("^tsi¬an;m^t", "^tyear^t")
Call WordGlossaryEngine("^tm;˚ø^t", "^tall^t")
Call WordGlossaryEngine("^tsw;lt;n^t", "^tnet^t")
Call WordGlossaryEngine("^tmen^t", "^tfather^t")
Call WordGlossaryEngine("^t÷;køiya÷qø^t", "^tgr.gr.grandparent^t")
Call WordGlossaryEngine("^tm;∫;^t", "^tchild^t")
Call WordGlossaryEngine("^tsi¬;^t", "^tgrandparent^t")
Call WordGlossaryEngine("^tsta¬;s^t", "^tspouse^t")
Call WordGlossaryEngine("^t;∫√e÷^t", "^tolder^t")
Call WordGlossaryEngine("^tten^t", "^tmother^t")
Call WordGlossaryEngine("^tkøan^t", "^tborn^t")
Call WordGlossaryEngine("^tsqe÷;q^t", "^tyounger.sibling^t")
Call WordGlossaryEngine("^tmeµ;∫;^t", "^tchildren^t")
Call WordGlossaryEngine("^tye¥s;¬;^t", "^ttwo people^t")
Call WordGlossaryEngine("^tßqøal;w;n^t", "^tthoughts^t")
Call WordGlossaryEngine("^ts®eni÷^t", "^twoman^t")
Call WordGlossaryEngine("^ts®;n®eni÷^t", "^twomen^t")
Call WordGlossaryEngine("^tna∫;ça÷^t", "^tone.person^t")
Call WordGlossaryEngine("^tß;y;®^t", "^tolder.sibling^t")
Call WordGlossaryEngine("^tl;plit^t", "^tpriest^t")
Call WordGlossaryEngine("^ttint;n^t", "^tbell^t")
Call WordGlossaryEngine("^t÷i¸;m^t", "^tget.dressed^t")
Call WordGlossaryEngine("^tœe¬;mi÷^t", "^tgirls^t")
Call WordGlossaryEngine("^tsya®^t", "^tfirewood^t")
Call WordGlossaryEngine("^tsπeœ;m^t", "^tflower^t")
Call WordGlossaryEngine("^t≈;÷aƒ;n^t", "^tfour^t")
Call WordGlossaryEngine("^t®ew;†^t", "^therring^t")
Call WordGlossaryEngine("^tsw;lt;n^t", "^tnet^t")
Call WordGlossaryEngine("^tƒi^t", "^tbig^t")
Call WordGlossaryEngine("^t∫an^t", "^tvery^t")
Call WordGlossaryEngine("^tl;plaß^t", "^tboard^t")
Call WordGlossaryEngine("^txø;t;s^t", "^theavy^t")
Call WordGlossaryEngine("^tsw;¥qe÷^t", "^tman^t")
Call WordGlossaryEngine("^ts˚øey^t", "^tcannot^t")
Call WordGlossaryEngine("^tn;çiµ^t", "^twhy^t")
Call WordGlossaryEngine("^t√e÷^t", "^talso^t")
Call WordGlossaryEngine("^tkøe:ƒ^t", "^tthen^t")
Call WordGlossaryEngine("^tkøe÷;®^t", "^tthen^t")
Call WordGlossaryEngine("^tƒ;®^t", "^tindeed^t")
Call WordGlossaryEngine("^ttxø^t", "^tonly^t")
Call WordGlossaryEngine("^tsn;xø;®^t", "^tcanoe^t")
Call WordGlossaryEngine("^t÷ay;m^t", "^tslow^t")
Call WordGlossaryEngine("^ty;_e∫^t", "^tfirst^t")
Call WordGlossaryEngine("^tswa_l;s^t", "^tboys^t")
Call WordGlossaryEngine("^tßxø;_weli^t", "^tparents^t")
Call WordGlossaryEngine("^ts;∫i_^t", "^tin^t")
Call WordGlossaryEngine("^tcec;_^t", "^tbeach^t")
Call WordGlossaryEngine("^ts;_œ^t", "^tfind^t")
Call WordGlossaryEngine("^tneç;_;c^t", "^tone.hundred^t")
Call WordGlossaryEngine("^t;_w;¥qe÷^t", "^tmen^t")
Call WordGlossaryEngine("^tƒiƒ;^t", "^tbig.PL^t")


' Restore hyphens
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t-^t"
        .Replacement.Text = "-"
        .Forward = False
        .Wrap = wdFindContinue
        .Format = True
     End With
    Selection.Find.Execute Replace:=wdReplaceAll

' Change glossed forms to auto/black
    Selection.Find.ClearFormatting
    Selection.Find.Font.Color = wdColorTeal
    Selection.Find.Font.Name = "Times"
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = wdColorAutomatic
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' Change untabbed line back to black
    Selection.Find.ClearFormatting
    Selection.Find.Font.Color = wdColorGray50
    Selection.Find.Font.Name = "Straight"
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = wdColorAutomatic
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' Remove tabs at beginning of line
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "#^t"
        .Replacement.Text = ""
'        .Forward = False
        .Wrap = wdFindContinue
        .Format = True
     End With
    Selection.Find.Execute Replace:=wdReplaceAll

 ' Remove tabs at end of line
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t#"
        .Replacement.Text = ""
'        .Forward = False
        .Wrap = wdFindContinue
        .Format = True
     End With
    Selection.Find.Execute Replace:=wdReplaceAll
    MsgBox ("I'm done.")


End Sub
Sub WordGlossaryEngine(FindText As String, ReplaceText As String)
Attribute WordGlossaryEngine.VB_Description = "Macro created 8/25/09 by Zoey Peterson"
Attribute WordGlossaryEngine.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.GramMorphList"

With Selection.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = FindText
    .Replacement.Text = ReplaceText
    
    Selection.Find.Font.Color = wdColorTeal
'    .Forward = True
    .Wrap = wdFindContinue
'    .Format = True
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False

    Selection.Find.Replacement.Font.Name = "Times"
'    Selection.Find.Replacement.Font.Color = wdAuto
    Selection.Find.Replacement.Font.SmallCaps = False
'    Selection.Find.Replacement.Font.AllCaps = False

'    Do While .Execute
        'Keep going until nothing found
        .Execute Replace:=wdReplaceAll
'    Loop
    'Free up some memory
'    ActiveDocument.UndoClear
End With

End Sub
Sub GramMorphEngine(FindText As String, ReplaceText As String)
'
' GramMorphEngine Macro
' Macro created 8/25/09 by Zoey Peterson

With Selection.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = FindText
    .Replacement.Text = ReplaceText
    
    Selection.Find.Font.Color = wdColorTeal
'    .Forward = True
    .Wrap = wdFindContinue
'    .Format = True
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False

    Selection.Find.Replacement.Font.Name = "Times"
'    Selection.Find.Replacement.Font.Color = wdAuto
    Selection.Find.Replacement.Font.SmallCaps = True
    Selection.Find.Replacement.Font.AllCaps = False

'    Do While .Execute
        'Keep going until nothing found
        .Execute Replace:=wdReplaceAll
'    Loop
    'Free up some memory
'    ActiveDocument.UndoClear
End With

End Sub
