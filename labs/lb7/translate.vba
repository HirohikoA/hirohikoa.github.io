Attribute VB_Name = "Module3"
Sub Раскладка()
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
    .Execute "q", ReplaceWith:="й", Replace:=wdReplaceAll
    .Execute "w", ReplaceWith:="ц", Replace:=wdReplaceAll
    .Execute "e", ReplaceWith:="у", Replace:=wdReplaceAll
    .Execute "r", ReplaceWith:="к", Replace:=wdReplaceAll
    .Execute "t", ReplaceWith:="е", Replace:=wdReplaceAll
    .Execute "y", ReplaceWith:="н", Replace:=wdReplaceAll
    .Execute "u", ReplaceWith:="г", Replace:=wdReplaceAll
    .Execute "i", ReplaceWith:="ш", Replace:=wdReplaceAll
    .Execute "o", ReplaceWith:="щ", Replace:=wdReplaceAll
    .Execute "p", ReplaceWith:="з", Replace:=wdReplaceAll
    .Execute "[", ReplaceWith:="х", Replace:=wdReplaceAll
    .Execute "]", ReplaceWith:="ъ", Replace:=wdReplaceAll
    .Execute "a", ReplaceWith:="ф", Replace:=wdReplaceAll
    .Execute "s", ReplaceWith:="ы", Replace:=wdReplaceAll
    .Execute "d", ReplaceWith:="в", Replace:=wdReplaceAll
    .Execute "f", ReplaceWith:="а", Replace:=wdReplaceAll
    .Execute "g", ReplaceWith:="п", Replace:=wdReplaceAll
    .Execute "h", ReplaceWith:="р", Replace:=wdReplaceAll
    .Execute "j", ReplaceWith:="о", Replace:=wdReplaceAll
    .Execute "k", ReplaceWith:="л", Replace:=wdReplaceAll
    .Execute "l", ReplaceWith:="д", Replace:=wdReplaceAll
    .Execute ";", ReplaceWith:="ж", Replace:=wdReplaceAll
    .Execute "'", ReplaceWith:="э", Replace:=wdReplaceAll
    .Execute "z", ReplaceWith:="я", Replace:=wdReplaceAll
    .Execute "x", ReplaceWith:="ч", Replace:=wdReplaceAll
    .Execute "c", ReplaceWith:="с", Replace:=wdReplaceAll
    .Execute "v", ReplaceWith:="м", Replace:=wdReplaceAll
    .Execute "b", ReplaceWith:="и", Replace:=wdReplaceAll
    .Execute "n", ReplaceWith:="т", Replace:=wdReplaceAll
    .Execute "m", ReplaceWith:="ь", Replace:=wdReplaceAll
    .Execute ",", ReplaceWith:="б", Replace:=wdReplaceAll
    .Execute ".", ReplaceWith:="ю", Replace:=wdReplaceAll
    .Execute "/", ReplaceWith:=".", Replace:=wdReplaceAll
    .Execute "?", ReplaceWith:=",", Replace:=wdReplaceAll
    .Execute "<", ReplaceWith:="Б", Replace:=wdReplaceAll
    .Execute ">", ReplaceWith:="Ю", Replace:=wdReplaceAll
    .Execute "{", ReplaceWith:="Х", Replace:=wdReplaceAll
    .Execute "}", ReplaceWith:="Ъ", Replace:=wdReplaceAll
    .Execute ":", ReplaceWith:="Ж", Replace:=wdReplaceAll
End With
End Sub
