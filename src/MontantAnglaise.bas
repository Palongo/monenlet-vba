Option Explicit

'==============================================================================
' MODULE  : MontantAnglaise
' AUTEUR  : Justin FARALAHY / MAAS
' DATE    : 18/04/2026
' VERSION : 1.1 -- Bugs critiques corriges (v1.0)
'
' FONCTION PRINCIPALE : MONENLET_EN(Valeur, [NbDecimales], [Devise])
'
' EXEMPLES VALIDES :
'   MONENLET_EN(1250)
'   -> "ONE THOUSAND TWO HUNDRED AND FIFTY ARIARYS"
'
'   MONENLET_EN(1250.25; 2; "ARIARY")
'   -> "ONE THOUSAND TWO HUNDRED AND FIFTY ARIARYS AND TWENTY-FIVE CENTS"
'
'   MONENLET_EN(1001)    -> "ONE THOUSAND AND ONE ARIARYS"
'   MONENLET_EN(1100)    -> "ONE THOUSAND ONE HUNDRED ARIARYS"
'   MONENLET_EN(200)     -> "TWO HUNDRED ARIARYS"
'   MONENLET_EN(1000000) -> "ONE MILLION ARIARYS"
'   MONENLET_EN(2000000) -> "TWO MILLION ARIARYS"
'
' REGLES LINGUISTIQUES ANGLAISES (British standard) :
'   - Hyphen between tens and units : TWENTY-ONE, FORTY-FIVE
'   - "AND" inserted after HUNDRED when a remainder follows
'     ex: 101 -> ONE HUNDRED AND ONE
'   - "AND" inserted after THOUSAND only when remainder < 100
'     ex: 1001 -> ONE THOUSAND AND ONE
'     ex: 1021 -> ONE THOUSAND AND TWENTY-ONE
'     ex: 1100 -> ONE THOUSAND ONE HUNDRED  (no AND, remainder >= 100)
'   - "ONE THOUSAND" (unlike French "MILLE" which drops "UN")
'   - HUNDRED / THOUSAND / MILLION / BILLION never take a plural S
'     ex: "TWO HUNDRED", "TWO MILLION"  (differ de FR "DEUX CENTS", "DEUX MILLIONS")
'   - No special cases for 70-79 / 80-99
'     (unlike French SOIXANTE-DIX, QUATRE-VINGT)
'
' DIFFERENCES vs MONENLET_FR :
'   FR: SOIXANTE ET ONZE (71)  -> EN: SEVENTY-ONE
'   FR: QUATRE-VINGTS (80)     -> EN: EIGHTY
'   FR: MILLE (1000, no UN)    -> EN: ONE THOUSAND
'   FR: DEUX CENTS (S on 200)  -> EN: TWO HUNDRED (no S ever)
'   FR: DEUX MILLIONS (S)      -> EN: TWO MILLION (no S)
'   FR: ET (21,31...)          -> EN: AND (only after HUNDRED/THOUSAND + small rest)
'
' CORRECTIFS v1.1 (vs v1.0) :
'   BUG 1 -- "Currency" est un mot-cle reserve VBA (type monetaire natif).
'            Utilise comme nom de parametre, il provoquait une erreur de
'            compilation : toutes les lignes de la fonction s'affichaient
'            en rouge et MONENLET_EN n'apparaissait pas dans =MON...
'            Fix : renomme en "Devise" (coherence avec MONENLET_FR et MONENLET_MG)
'   BUG 2 -- Auto_Open en conflit avec les autres modules.
'            Trois "Auto_Open" dans le meme classeur -> "Ambiguous name
'            detected" -> projet non compilable -> aucune fonction visible.
'            Fix : renomme en RegisterMONENLET_EN.
'            Voir ModuleRegistration.bas pour le Auto_Open centralise.
'
' HISTORIQUE :
'   v1.0 -- Version initiale -- 2 bugs critiques
'   v1.1 -- BUG 1 et BUG 2 corriges, style mis a jour
'
' STRUCTURE IDENTIQUE aux modules :
'   MontantMalagasy.bas  -- MONENLET_MG
'   MontantFrancaise.bas -- MONENLET_FR
'==============================================================================


' ─────────────────────────────────────────────────────────────────────────────
' TABLEAUX DE NOMENCLATURE -- Niveau module (une seule initialisation)
' ─────────────────────────────────────────────────────────────────────────────

' Unites et irreguliers (0 a 19)
Private gUnits As Variant

' Dizaines (index 2 a 9 : TWENTY a NINETY)
Private gTens  As Variant

' Drapeau d'initialisation
Private gInit  As Boolean


Private Sub InitTableaux()
    '--------------------------------------------------------------------------
    ' Initialise les tableaux une seule fois par session.
    ' Appelee automatiquement avant chaque conversion.
    '--------------------------------------------------------------------------
    If gInit Then Exit Sub

    gUnits = Array("", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", _
                   "SEVEN", "EIGHT", "NINE", "TEN", "ELEVEN", "TWELVE", _
                   "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN", _
                   "SEVENTEEN", "EIGHTEEN", "NINETEEN")

    gTens = Array("", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", _
                  "SIXTY", "SEVENTY", "EIGHTY", "NINETY")

    gInit = True
End Sub


' ─────────────────────────────────────────────────────────────────────────────
' ENREGISTREMENT INFO-BULLE
' ─────────────────────────────────────────────────────────────────────────────

Public Sub RegisterMONENLET_EN()
    '--------------------------------------------------------------------------
    ' Enregistre MONENLET_EN aupres d'Excel via MacroOptions.
    '
    ' IMPORTANT : Ne pas renommer en "Auto_Open".
    '   Si les trois modules (MG, FR, EN) coexistent dans le meme classeur,
    '   trois "Auto_Open" -> erreur "Ambiguous name detected".
    '   Ce Sub est appele par le module ModuleRegistration.bas.
    '
    ' Effet visible :
    '   - "Insert Function" (Shift+F3) : description + info-bulles arguments
    '   - Saisie automatique Excel : categorie "Finances EN"
    '--------------------------------------------------------------------------
    Dim argDesc(2) As String
    argDesc(0) = "Numeric amount to convert (e.g. 1250.75). " & _
                 "Supports up to 999,999,999,999 (999 billion). " & _
                 "Negative values are automatically converted to positive."
    argDesc(1) = "[Optional] Number of decimal digits (default: 2). " & _
                 "Set to 0 to ignore cents."
    argDesc(2) = "[Optional] Currency label after the integer part " & _
                 "(default: ""ARIARY""). ""ARIARY"" auto-pluralises to " & _
                 """ARIARYS"". Any other value is displayed as-is."

    Application.MacroOptions _
        Macro:="MONENLET_EN", _
        Description:="Converts a numeric amount into English words " & _
                      "(British banking / legal standard). Handles AND " & _
                      "after HUNDRED/THOUSAND, hyphens in compound tens, " & _
                      "and correct plural for ARIARY. " & _
                      "Ex: MONENLET_EN(1250.25) -> " & _
                      """ONE THOUSAND TWO HUNDRED AND FIFTY ARIARYS " & _
                      "AND TWENTY-FIVE CENTS"".", _
        Category:="Finances EN", _
        ArgumentDescriptions:=argDesc
End Sub


' ─────────────────────────────────────────────────────────────────────────────
' FONCTION PUBLIQUE -- Point d'entree  (nom court pour usage Excel)
' ─────────────────────────────────────────────────────────────────────────────

Public Function MONENLET_EN( _
        ByVal Valeur             As Double, _
        Optional ByVal NbDecimales As Integer = 2, _
        Optional ByVal Devise      As String  = "ARIARY") As String
    '--------------------------------------------------------------------------
    ' Converts a numeric amount into English words (British banking standard).
    '
    ' Syntaxe    : MONENLET_EN(Valeur; [NbDecimales]; [Devise])
    ' Parametres :
    '   Valeur       -- montant a convertir (negatif -> converti en Abs)
    '   NbDecimales  -- precision decimale, defaut 2 (min 0)
    '   Devise       -- libelle de devise, defaut "ARIARY"
    '                   NOTE : nomme "Devise" (et non "Currency" qui est un
    '                   type de donnee reserve VBA -> erreur de compilation)
    '
    ' Retourne : chaine en majuscules selon les normes bancaires britanniques.
    '--------------------------------------------------------------------------

    Call InitTableaux

    If NbDecimales < 0 Then NbDecimales = 2

    Valeur = Abs(Valeur)

    Dim entier As Double
    Dim cents  As Long

    entier = Fix(Valeur)

    ' CLng(Round()) evite la perte de precision flottante sur les centimes
    If NbDecimales = 0 Then
        cents = 0
    Else
        cents = CLng(Round((Valeur - entier) * (10 ^ NbDecimales), 0))
    End If

    ' Partie entiere + devise
    Dim texte As String
    texte = NombreEnLettresEN(entier)

    If UCase(Devise) = "ARIARY" Then
        If entier > 1 Then
            texte = texte & " ARIARYS"
        Else
            texte = texte & " ARIARY"
        End If
    Else
        texte = texte & " " & Devise
    End If

    ' Partie decimale (cents)
    If NbDecimales > 0 And cents > 0 Then
        texte = texte & " AND " & NombreEnLettresEN(cents) & " CENT"
        If cents > 1 Then texte = texte & "S"
    End If

    MONENLET_EN = Application.WorksheetFunction.Trim(texte)
End Function


' ─────────────────────────────────────────────────────────────────────────────
' CONVERSION RECURSIVE -- Nombre entier -> lettres anglaises
' ─────────────────────────────────────────────────────────────────────────────

Private Function NombreEnLettresEN(ByVal N As Double) As String
    '--------------------------------------------------------------------------
    ' Converts an integer recursively into English words (British standard).
    '
    ' Plages traitees :
    '   0              -> ZERO
    '   1-19           -> direct lookup (gUnits)
    '   20-99          -> gTens + optional hyphen + gUnits
    '   100-999        -> X HUNDRED [AND rest]
    '   1 000-999 999  -> X THOUSAND [AND/space rest]
    '   1 M-999 M      -> X MILLION [rest]   (no S)
    '   1 B-999 B      -> X BILLION [rest]   (no S)
    '   Else           -> "#NUMBER TOO LARGE"
    '
    ' Variables locales declarees en tete (pas dans les Case) pour eviter
    ' tout probleme de compilation avec VBA en mode recursif.
    '--------------------------------------------------------------------------

    Dim texte As String
    Dim u99   As Integer
    Dim r100  As Integer
    Dim r1k   As Integer
    Dim nbMil As Double
    Dim rMil  As Double
    Dim nbBil As Double
    Dim rBil  As Double

    Select Case N

        ' -- Zero ─────────────────────────────────────────────────────────────
        Case 0
            texte = "ZERO"

        ' -- 1 a 19 : tableau direct ──────────────────────────────────────────
        Case 1 To 19
            texte = gUnits(CInt(N))

        ' -- 20 a 99 : dizaines + tiret ───────────────────────────────────────
        '   Hyphen entre dizaine et unite : TWENTY-ONE, NINETY-NINE
        '   Pas de connecteur "ET" comme en francais
        Case 20 To 99
            texte = gTens(CInt(Int(N / 10)))
            u99 = CInt(N Mod 10)
            If u99 > 0 Then texte = texte & "-" & gUnits(u99)

        ' -- 100 a 999 : centaines ────────────────────────────────────────────
        '   HUNDRED ne prend jamais de S (differ de FR CENTS)
        '   "AND" toujours insere avant le reste si > 0
        Case 100 To 999
            texte = gUnits(CInt(Int(N / 100))) & " HUNDRED"
            r100 = CInt(N Mod 100)
            If r100 > 0 Then texte = texte & " AND " & NombreEnLettresEN(r100)

        ' -- 1 000 a 999 999 : milliers ───────────────────────────────────────
        '   "ONE THOUSAND" : ONE obligatoire (differ de FR "MILLE" sans "UN")
        '   "AND" apres THOUSAND uniquement si reste < 100 :
        '     1001 -> ONE THOUSAND AND ONE
        '     1100 -> ONE THOUSAND ONE HUNDRED  (pas de AND)
        Case 1000 To 999999
            texte = NombreEnLettresEN(Int(N / 1000)) & " THOUSAND"
            r1k = CInt(N Mod 1000)
            If r1k > 0 Then
                If r1k < 100 Then
                    texte = texte & " AND " & NombreEnLettresEN(r1k)
                Else
                    texte = texte & " " & NombreEnLettresEN(r1k)
                End If
            End If

        ' -- 1 000 000 a 999 999 999 : millions ───────────────────────────────
        '   Pas de S sur MILLION ("TWO MILLION", pas "TWO MILLIONS")
        Case 1000000 To 999999999
            nbMil = Int(N / 1000000#)
            texte = NombreEnLettresEN(nbMil) & " MILLION"
            rMil = N - nbMil * 1000000#
            If rMil > 0 Then texte = texte & " " & NombreEnLettresEN(rMil)

        ' -- 1 000 000 000 a 999 999 999 999 : milliards ──────────────────────
        '   Pas de S sur BILLION
        '   Reste par soustraction (pas Mod) -> evite imprecision Double > 2^31
        Case 1000000000 To 999999999999#
            nbBil = Int(N / 1000000000#)
            texte = NombreEnLettresEN(nbBil) & " BILLION"
            rBil = N - nbBil * 1000000000#
            If rBil > 0 Then texte = texte & " " & NombreEnLettresEN(rBil)

        ' -- Depassement de capacite ──────────────────────────────────────────
        Case Else
            texte = "#NUMBER TOO LARGE"

    End Select

    NombreEnLettresEN = texte
End Function


' ─────────────────────────────────────────────────────────────────────────────
' PROCEDURE DE TEST (F5 sur Sub TestMontantEN dans l'IDE VBA)
' ─────────────────────────────────────────────────────────────────────────────

Public Sub TestMontantEN()
    Call InitTableaux

    Dim sep As String
    sep = String(72, "-")

    Debug.Print sep
    Debug.Print "TEST -- MONENLET_EN v1.1"
    Debug.Print sep

    Debug.Print ""
    Debug.Print "[AND RULE]"
    Debug.Print "101         -> " & MONENLET_EN(101, 0)
    Debug.Print "110         -> " & MONENLET_EN(110, 0)
    Debug.Print "1001        -> " & MONENLET_EN(1001, 0)
    Debug.Print "1021        -> " & MONENLET_EN(1021, 0)
    Debug.Print "1100        -> " & MONENLET_EN(1100, 0)
    Debug.Print "1101        -> " & MONENLET_EN(1101, 0)

    Debug.Print ""
    Debug.Print "[HYPHEN RULE]"
    Debug.Print "21          -> " & MONENLET_EN(21, 0)
    Debug.Print "71          -> " & MONENLET_EN(71, 0)
    Debug.Print "99          -> " & MONENLET_EN(99, 0)

    Debug.Print ""
    Debug.Print "[NO PLURAL ON MULTIPLIERS]"
    Debug.Print "200         -> " & MONENLET_EN(200, 0)
    Debug.Print "2000        -> " & MONENLET_EN(2000, 0)
    Debug.Print "2000000     -> " & MONENLET_EN(2000000, 0)
    Debug.Print "2000000000  -> " & MONENLET_EN(2000000000#, 0)

    Debug.Print ""
    Debug.Print "[ONE BEFORE THOUSAND/MILLION/BILLION]"
    Debug.Print "1000        -> " & MONENLET_EN(1000, 0)
    Debug.Print "1000000     -> " & MONENLET_EN(1000000, 0)
    Debug.Print "1000000000  -> " & MONENLET_EN(1000000000#, 0)

    Debug.Print ""
    Debug.Print "[DECIMALS]"
    Debug.Print "1250        -> " & MONENLET_EN(1250)
    Debug.Print "1250.25     -> " & MONENLET_EN(1250.25)
    Debug.Print "1250.01     -> " & MONENLET_EN(1250.01)
    Debug.Print "0.50        -> " & MONENLET_EN(0.5)

    Debug.Print ""
    Debug.Print "[EDGE CASES]"
    Debug.Print "0           -> " & MONENLET_EN(0)
    Debug.Print "1           -> " & MONENLET_EN(1, 0)
    Debug.Print "-500        -> " & MONENLET_EN(-500, 0)
    Debug.Print "100000      -> " & MONENLET_EN(100000, 0)
    Debug.Print "1352689.60  -> " & MONENLET_EN(1352689.6)
    Debug.Print "2245710096  -> " & MONENLET_EN(2245710096.15)
    Debug.Print "999999999999 -> " & MONENLET_EN(999999999999#, 0)
    Debug.Print sep
End Sub
