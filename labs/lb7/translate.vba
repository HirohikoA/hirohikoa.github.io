Attribute VB_Name = "Module3"
Sub ���������()
With ActiveDocument.Range.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute "q", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "w", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "e", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "r", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "t", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "y", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "u", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "i", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "o", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "p", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "[", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "]", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "a", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "s", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "d", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "f", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "g", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "h", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "j", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "k", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "l", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute ";", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "'", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "z", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "x", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "c", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "v", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "b", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "n", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "m", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute ",", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute ".", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "/", ReplaceWith:=".", Replace:=wdReplaceAll
    .Execute "?", ReplaceWith:=",", Replace:=wdReplaceAll
    .Execute "<", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute ">", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "{", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute "}", ReplaceWith:="�", Replace:=wdReplaceAll
    .Execute ":", ReplaceWith:="�", Replace:=wdReplaceAll
End With
End Sub
