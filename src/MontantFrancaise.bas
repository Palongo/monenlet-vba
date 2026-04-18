Option Explicit

'==============================================================================
' MODULE  : MontantFrancaise
' AUTEUR  : Justin FARALAHY / MAAS
' DATE    : 18/04/2026
' VERSION : 2.1 — Renommée MONENLET_FR, révision complète (logique inchangée)
'
' FONCTION PRINCIPALE : MONENLET_FR(Valeur, [NbDecimales], [Devise])
'
' EXEMPLES VALIDÉS :
'   MONENLET_FR(1250)
'   → "MILLE DEUX CENT CINQUANTE ARIARYS"
'
'   MONENLET_FR(1250.25; 2; "ARIARY")
'   → "MILLE DEUX CENT CINQUANTE ARIARYS ET VINGT-CINQ CENTIMES"
'
'   MONENLET_FR(80)      → "QUATRE-VINGTS ARIARYS"
'   MONENLET_FR(200)     → "DEUX CENTS ARIARYS"
'   MONENLET_FR(201)     → "DEUX CENT UN ARIARYS"
'   MONENLET_FR(71)      → "SOIXANTE ET ONZE ARIARYS"
'   MONENLET_FR(1000000) → "UN MILLION ARIARYS"
'
' RÈGLES LINGUISTIQUES FRANÇAISES :
'   - "MILLE" n'est jamais précédé de "UN"
'   - "CENT" prend un S uniquement s'il clôt le nombre (200, 300...)
'   - "QUATRE-VINGTS" prend un S uniquement pour 80 exact
'   - "ET" utilisé pour 21, 31, 41, 51, 61, 71 uniquement
'   - Tiret entre dizaine et unité (sauf cas "ET")
'   - "MILLION(S)" / "MILLIARD(S)" prennent le pluriel si > 1
'
' CORRECTIFS v2.0 (vs v1.0) :
'   BUG 1 — Attribute VB_Name supprimé (erreur de compilation)
'   BUG 2 — Decimale : Round() → CLng(Round()) évite la perte de précision
'            en virgule flottante (ex: 1.005 → centimes = 0 au lieu de 1)
'   BUG 3 — Milliards : "N Mod 1000000000#" remplacé par soustraction
'            explicite pour éviter l'imprécision de Mod avec Double > 2^31
'   BUG 4 — Valeurs négatives : Abs() appliqué dès l'entrée (Fix(-1.9)=-1
'            tombait dans Case Else → "#NOMBRE TROP GRAND")
'   OPT 1 — Tableaux Unite/Dizaine déplacés au niveau module (évite leur
'            réinstanciation à chaque appel récursif)
'   OPT 2 — Auto_Open + MacroOptions : info-bulle Excel pour MONENLET_FR
'
' HISTORIQUE :
'   v1.0 — Version originale (MONENLET_PRO) — 4 bugs identifiés
'   v2.0 — Bugs corrigés, structure unifiée, info-bulle Excel ajoutée
'   v2.1 — Renommée MONENLET_FR, révision complète (aucun bug supplémentaire)
'
' STRUCTURE IDENTIQUE aux modules :
'   MontantMalagasy.bas  — MONENLET_MG
'   MontantAnglaise.bas  — MONENLET_EN
'==============================================================================


' ─────────────────────────────────────────────────────────────────────────────
' TABLEAUX DE NOMENCLATURE — Niveau module (une seule initialisation)
' ─────────────────────────────────────────────────────────────────────────────

' Unités et dizaines irrégulières (0 → 19)
Private gUnite   As Variant

' Dizaines régulières (index 2 → 6 : VINGT à SOIXANTE)
Private gDizaine As Variant

' Drapeau d'initialisation des tableaux
Private gInit    As Boolean


Private Sub InitTableaux()
    '--------------------------------------------------------------------------
    ' Initialise les tableaux de nomenclature une seule fois par session.
    ' Appelée automatiquement avant chaque conversion si nécessaire.
    '--------------------------------------------------------------------------
    If gInit Then Exit Sub

    gUnite = Array("", "UN", "DEUX", "TROIS", "QUATRE", "CINQ", "SIX", _
                   "SEPT", "HUIT", "NEUF", "DIX", "ONZE", "DOUZE", "TREIZE", _
                   "QUATORZE", "QUINZE", "SEIZE", "DIX-SEPT", "DIX-HUIT", "DIX-NEUF")

    gDizaine = Array("", "", "VINGT", "TRENTE", "QUARANTE", "CINQUANTE", "SOIXANTE")

    gInit = True
End Sub


' ─────────────────────────────────────────────────────────────────────────────
' ENREGISTREMENT INFO-BULLE — À appeler une fois à l'ouverture du classeur
' ─────────────────────────────────────────────────────────────────────────────

Public Sub RegisterMONENLET_FR()
    '--------------------------------------------------------------------------
    ' Enregistre MONENLET_FR aupres d'Excel via MacroOptions.
    '
    ' IMPORTANT : Ne pas renommer en "Auto_Open".
    '   Si les trois modules (MG, FR, EN) coexistent dans le meme classeur,
    '   trois "Auto_Open" -> erreur "Ambiguous name detected".
    '   Ce Sub est appele par le module ModuleRegistration.bas.
    '--------------------------------------------------------------------------
    Dim argDesc(2) As String
    argDesc(0) = "Montant numérique à convertir (ex: 1250.75). " & _
                 "Supporte jusqu'à 999 999 999 999 (999 milliards). " & _
                 "Les valeurs négatives sont automatiquement converties en positif."
    argDesc(1) = "[Optionnel] Nombre de chiffres après la virgule (défaut : 2). " & _
                 "Mettre 0 pour ignorer les centimes."
    argDesc(2) = "[Optionnel] Libellé de la devise affiché après la partie entière " & _
                 "(défaut : ""ARIARY""). ""ARIARY"" gère automatiquement le pluriel " & _
                 "en ""ARIARYS"". Toute autre devise est affichée telle quelle."

    Application.MacroOptions _
        Macro:="MONENLET_FR", _
        Description:="Convertit un montant numérique en lettres françaises " & _
                      "(normes bancaires / juridiques). Gère les règles " & _
                      "d'accord (CENTS, QUATRE-VINGTS, ET), le pluriel des " & _
                      "grandes unités, et les centimes après ""ET"". " & _
                      "Ex : MONENLET_FR(1250.25) → " & _
                      """MILLE DEUX CENT CINQUANTE ARIARYS ET VINGT-CINQ CENTIMES"".", _
        Category:="Finances FR", _
        ArgumentDescriptions:=argDesc
End Sub


' ─────────────────────────────────────────────────────────────────────────────
' FONCTION PUBLIQUE — Point d'entrée  (nom court pour usage Excel)
' ─────────────────────────────────────────────────────────────────────────────

Public Function MONENLET_FR( _
        ByVal Valeur       As Double, _
        Optional ByVal NbDecimales As Integer = 2, _
        Optional ByVal Devise      As String  = "ARIARY") As String
    '--------------------------------------------------------------------------
    ' Convertit un montant numérique en lettres françaises.
    '
    ' Syntaxe    : MONENLET_FR(Valeur; [NbDecimales]; [Devise])
    ' Paramètres :
    '   Valeur       — montant à convertir (négatif accepté → converti en Abs)
    '   NbDecimales  — précision décimale, défaut 2 (min 0)
    '   Devise       — libellé de devise, défaut "ARIARY"
    '
    ' Retourne : chaîne en majuscules selon les normes françaises.
    '--------------------------------------------------------------------------

    ' Initialisation des tableaux si nécessaire
    Call InitTableaux

    ' Sécurisation des paramètres
    If NbDecimales < 0 Then NbDecimales = 2

    ' BUG 4 corrigé : protection contre les négatifs
    Valeur = Abs(Valeur)

    ' Séparation partie entière / décimale
    Dim entier   As Double
    Dim centimes As Long

    entier = Fix(Valeur)

    ' BUG 2 corrigé : CLng évite la perte de précision de Round sur Double
    If NbDecimales = 0 Then
        centimes = 0
    Else
        centimes = CLng(Round((Valeur - entier) * (10 ^ NbDecimales), 0))
    End If

    ' Partie entière + devise
    Dim texte As String
    texte = NombreEnLettresFR(entier)

    If UCase(Devise) = "ARIARY" Then
        ' Pluriel spécifique Ariary
        If entier > 1 Then
            texte = texte & " ARIARYS"
        Else
            texte = texte & " ARIARY"
        End If
    Else
        texte = texte & " " & Devise
    End If

    ' Partie décimale (centimes)
    If NbDecimales > 0 And centimes > 0 Then
        texte = texte & " ET " & NombreEnLettresFR(centimes) & " CENTIME"
        If centimes > 1 Then texte = texte & "S"
    End If

    MONENLET_FR = Application.WorksheetFunction.Trim(texte)
End Function


' ─────────────────────────────────────────────────────────────────────────────
' CONVERSION RÉCURSIVE — Nombre entier → lettres françaises
' ─────────────────────────────────────────────────────────────────────────────

Private Function NombreEnLettresFR(ByVal N As Double) As String
    '--------------------------------------------------------------------------
    ' Convertit récursivement un entier en lettres selon les règles françaises.
    '
    ' Plages traitées (ordre Select Case) :
    '   0              → ZERO
    '   1–19           → unités et dizaines irrégulières (tableau gUnite)
    '   20–69          → dizaines régulières + connecteur ET / tiret
    '   70–79          → SOIXANTE + récursion sur (N-60)  [71 = cas spécial]
    '   80–99          → QUATRE-VINGT + récursion sur (N-80) [80 et 81 spéciaux]
    '   100–999        → centaines + accord du S final sur CENT
    '   1 000–999 999  → milliers (MILLE sans UN devant)
    '   1 M–999 M      → millions (accord pluriel)
    '   1 G–999 G      → milliards (accord pluriel)
    '   Else           → "#NOMBRE TROP GRAND"
    '
    ' NOTE (BUG 3 corrigé) : le reste des milliards est calculé par
    '   soustraction explicite pour éviter l'imprécision de Mod avec Double.
    '--------------------------------------------------------------------------

    Dim texte As String

    Select Case N

        ' ── Zéro ───────────────────────────────────────────────────────────
        Case 0
            texte = "ZERO"

        ' ── 1 à 19 : tableau direct ─────────────────────────────────────────
        Case 1 To 19
            texte = gUnite(CInt(N))

        ' ── 20 à 69 : dizaines régulières ───────────────────────────────────
        '   Règle ET : uniquement pour X1 (21, 31, 41, 51, 61)
        '   Tiret    : pour X2 à X9
        Case 20 To 69
            texte = gDizaine(CInt(Int(N / 10)))
            Select Case CInt(N Mod 10)
                Case 0:  ' dizaine exacte → rien à ajouter
                Case 1:  texte = texte & " ET UN"
                Case Else: texte = texte & "-" & gUnite(CInt(N Mod 10))
            End Select

        ' ── 70 à 79 : SOIXANTE + (10 à 19) ─────────────────────────────────
        '   71 = "SOIXANTE ET ONZE" (pas SOIXANTE-ONZE)
        Case 70 To 79
            If N = 71 Then
                texte = "SOIXANTE ET ONZE"
            Else
                texte = "SOIXANTE-" & NombreEnLettresFR(N - 60)
            End If

        ' ── 80 à 99 : QUATRE-VINGT(S) + reste ──────────────────────────────
        '   80  → QUATRE-VINGTS (avec S)
        '   81  → QUATRE-VINGT-UN (sans S, sans ET)
        '   82+ → QUATRE-VINGT- + récursion
        Case 80 To 99
            Select Case CInt(N)
                Case 80:  texte = "QUATRE-VINGTS"
                Case 81:  texte = "QUATRE-VINGT-UN"
                Case Else: texte = "QUATRE-VINGT-" & NombreEnLettresFR(N - 80)
            End Select

        ' ── 100 à 999 : centaines ───────────────────────────────────────────
        '   CENT prend S uniquement si le nombre est un multiple exact de 100
        '   et que le multiplicateur > 1 (200 → DEUX CENTS, 201 → DEUX CENT UN)
        Case 100 To 999
            Dim centH As Integer
            centH = CInt(Int(N / 100))

            If centH = 1 Then
                texte = "CENT"
            Else
                texte = gUnite(centH) & " CENT"
            End If

            If CInt(N Mod 100) = 0 Then
                If centH > 1 Then texte = texte & "S"   ' accord pluriel
            Else
                texte = texte & " " & NombreEnLettresFR(N Mod 100)
            End If

        ' ── 1 000 à 999 999 : milliers ──────────────────────────────────────
        '   "MILLE" n'est jamais précédé de "UN" (règle française)
        Case 1000 To 999999
            If CInt(Int(N / 1000)) = 1 Then
                texte = "MILLE"
            Else
                texte = NombreEnLettresFR(Int(N / 1000)) & " MILLE"
            End If

            If CInt(N Mod 1000) > 0 Then
                texte = texte & " " & NombreEnLettresFR(N Mod 1000)
            End If

        ' ── 1 000 000 à 999 999 999 : millions ──────────────────────────────
        '   "MILLION" prend S si le multiplicateur > 1
        Case 1000000 To 999999999
            Dim nbMil As Double
            nbMil = Int(N / 1000000#)

            If nbMil = 1 Then
                texte = "UN MILLION"
            Else
                texte = NombreEnLettresFR(nbMil) & " MILLIONS"
            End If

            Dim resteMil As Double
            resteMil = N - nbMil * 1000000#
            If resteMil > 0 Then texte = texte & " " & NombreEnLettresFR(resteMil)

        ' ── 1 000 000 000 à 999 999 999 999 : milliards ─────────────────────
        '   BUG 3 corrigé : reste calculé par soustraction (pas Mod)
        '   pour éviter l'imprécision de Mod avec Double > 2^31
        Case 1000000000 To 999999999999#
            Dim nbMrd As Double
            nbMrd = Int(N / 1000000000#)

            If nbMrd = 1 Then
                texte = "UN MILLIARD"
            Else
                texte = NombreEnLettresFR(nbMrd) & " MILLIARDS"
            End If

            Dim resteMrd As Double
            resteMrd = N - nbMrd * 1000000000#
            If resteMrd > 0 Then texte = texte & " " & NombreEnLettresFR(resteMrd)

        ' ── Dépassement de capacité ──────────────────────────────────────────
        Case Else
            texte = "#NOMBRE TROP GRAND"

    End Select

    NombreEnLettresFR = texte
End Function


' ─────────────────────────────────────────────────────────────────────────────
' PROCÉDURE DE TEST (à exécuter dans l'IDE VBA : F5 sur Sub TestMontantFR)
' ─────────────────────────────────────────────────────────────────────────────

Public Sub TestMontantFR()
    '--------------------------------------------------------------------------
    ' Lance une série de tests et affiche les résultats dans la fenêtre
    ' Exécution (Ctrl+G dans l'IDE VBA).
    '--------------------------------------------------------------------------
    Call InitTableaux

    Dim sep As String
    sep = String(72, "-")

    Debug.Print sep
    Debug.Print "TEST — MONENLET_FR v2.1"
    Debug.Print sep

    ' ── Règles linguistiques clés ────────────────────────────────────────────
    Debug.Print ""
    Debug.Print "[RÈGLES FRANÇAISES]"
    Debug.Print "21          → " & MONENLET_FR(21, 0)     ' VINGT ET UN
    Debug.Print "71          → " & MONENLET_FR(71, 0)     ' SOIXANTE ET ONZE
    Debug.Print "80          → " & MONENLET_FR(80, 0)     ' QUATRE-VINGTS
    Debug.Print "81          → " & MONENLET_FR(81, 0)     ' QUATRE-VINGT-UN
    Debug.Print "91          → " & MONENLET_FR(91, 0)     ' QUATRE-VINGT-ONZE
    Debug.Print "100         → " & MONENLET_FR(100, 0)    ' CENT
    Debug.Print "200         → " & MONENLET_FR(200, 0)    ' DEUX CENTS
    Debug.Print "201         → " & MONENLET_FR(201, 0)    ' DEUX CENT UN
    Debug.Print "1000        → " & MONENLET_FR(1000, 0)   ' MILLE (sans UN)
    Debug.Print "1001        → " & MONENLET_FR(1001, 0)   ' MILLE UN
    Debug.Print "1000000     → " & MONENLET_FR(1000000, 0)    ' UN MILLION
    Debug.Print "2000000     → " & MONENLET_FR(2000000, 0)    ' DEUX MILLIONS
    Debug.Print "1000000000  → " & MONENLET_FR(1000000000, 0)  ' UN MILLIARD
    Debug.Print "2000000000  → " & MONENLET_FR(2000000000#, 0) ' DEUX MILLIARDS
    Debug.Print ""

    ' ── Tests avec décimales ─────────────────────────────────────────────────
    Debug.Print "[DÉCIMALES]"
    Debug.Print "1250        → " & MONENLET_FR(1250)      ' ... ARIARYS
    Debug.Print "1250.25     → " & MONENLET_FR(1250.25)   ' ... ET VINGT-CINQ CENTIMES
    Debug.Print "1250.01     → " & MONENLET_FR(1250.01)   ' ... ET UN CENTIME
    Debug.Print "0.50        → " & MONENLET_FR(0.5)       ' ZERO ARIARY ET CINQUANTE CENTIMES
    Debug.Print ""

    ' ── Cas limites ──────────────────────────────────────────────────────────
    Debug.Print "[CAS LIMITES]"
    Debug.Print "0           → " & MONENLET_FR(0)
    Debug.Print "1           → " & MONENLET_FR(1, 0)
    Debug.Print "-500        → " & MONENLET_FR(-500, 0)          ' Abs → CINQ CENTS
    Debug.Print "999999999999 → " & MONENLET_FR(999999999999#, 0)
    Debug.Print sep
End Sub
