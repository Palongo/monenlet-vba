Option Explicit

'==============================================================================
' MODULE  : ModuleRegistration
' AUTEUR  : Justin FARALAHY / MAAS
' DATE    : 18/04/2026
' VERSION : 1.0
'
' ROLE :
'   Module centralise pour l'enregistrement des trois fonctions UDF
'   aupres d'Excel (info-bulles, categories, descriptions d'arguments).
'
'   PROBLEME resolu :
'   Chaque module (MontantMalagasy, MontantFrancaise, MontantAnglaise)
'   avait son propre "Auto_Open". Lorsque les trois modules coexistent
'   dans le meme classeur (.xlsm), VBA leve l'erreur :
'     "Ambiguous name detected: Auto_Open"
'   -> le projet entier ne compile pas -> aucune fonction n'apparait
'   dans la liste Excel (=MON...).
'
'   SOLUTION :
'   - Chaque module expose desormais un Sub d'enregistrement unique :
'       RegisterMONENLET_MG  (dans MontantMalagasy.bas)
'       RegisterMONENLET_FR  (dans MontantFrancaise.bas)
'       RegisterMONENLET_EN  (dans MontantAnglaise.bas)
'   - Ce module contient le SEUL Auto_Open du classeur.
'     Il appelle les trois Sub ci-dessus.
'
' UTILISATION :
'   1. Importer les quatre fichiers .bas dans le meme classeur .xlsm :
'        MontantMalagasy.bas
'        MontantFrancaise.bas
'        MontantAnglaise.bas
'        ModuleRegistration.bas   <- ce fichier
'   2. Sauvegarder en .xlsm (pas .xlsx).
'   3. A la prochaine ouverture du fichier, Auto_Open() s'execute
'      automatiquement et enregistre les trois fonctions.
'   4. En cas de besoin manuel : Alt+F8 -> "Auto_Open" -> Executer
'
' FONCTIONS ENREGISTREES :
'   MONENLET_MG  -- Montant en lettres Malgache   (categorie "Finances MG")
'   MONENLET_FR  -- Montant en lettres Francais   (categorie "Finances FR")
'   MONENLET_EN  -- Montant en lettres Anglais    (categorie "Finances EN")
'==============================================================================


Public Sub Auto_Open()
    '--------------------------------------------------------------------------
    ' Point d'entree unique pour l'enregistrement de toutes les UDF.
    ' Appele automatiquement par Excel a l'ouverture du classeur.
    '--------------------------------------------------------------------------

    ' Enregistrement Malgache
    Call RegisterMONENLET_MG

    ' Enregistrement Francais
    Call RegisterMONENLET_FR

    ' Enregistrement Anglais
    Call RegisterMONENLET_EN

End Sub
