Attribute VB_Name = "TraitementWord"
Sub ColorCrossReferences()
Attribute ColorCrossReferences.VB_Description = "Modifie le format des références croisées (Cross-References, en anglais) du document pour le style ""Emphase intense"""
Attribute ColorCrossReferences.VB_ProcData.VB_Invoke_Func = "Project.Traitement.ColorCrossReferences"
'   Macro qui modifie toutes les lien intra-document (Cross-References en anglais)
'   Afin de conserver la mise en forme spéciale
'   Source : https://superuser.com/questions/13531/is-it-possible-to-assign-a-specific-style-to-all-cross-references-in-word-2007
    ActiveDocument.ActiveWindow.View.ShowFieldCodes = True
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles( _
    "Emphase intense")
    With Selection.Find
        .Text = "^19 REF"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveDocument.ActiveWindow.View.ShowFieldCodes = False
    
    ' Convertir toutes les Cross-References avec le CHARFORMAT pour conserver le format lorsqu'on enregistre sous PDF
    Call SetCHARFORMAT
End Sub


Sub SetCHARFORMAT()
'
'   Set CHARFORMAT switch to all {REF} fields. Replace MERGEFORMAT.
'   Source : https://superuser.com/questions/13531/is-it-possible-to-assign-a-specific-style-to-all-cross-references-in-word-2007
'
    Dim oField As Field
    Dim oRng As Range
    For Each oField In ActiveDocument.Fields
        If InStr(1, oField.Code, "REF ") = 2 Then
            If InStr(1, oField.Code, "MERGEFORMAT") <> 0 Then
                oField.Code.Text = Replace(oField.Code.Text, "MERGEFORMAT", "CHARFORMAT")
            End If
            If InStr(1, oField.Code, "CHARFORMAT") = 0 Then
                oField.Code.Text = oField.Code.Text + "\* CHARFORMAT "
            End If
        End If
    Next oField
End Sub
