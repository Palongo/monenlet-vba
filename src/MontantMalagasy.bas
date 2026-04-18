Option Explicit

'==============================================================================
' MODULE  : MontantMalagasy
' AUTEUR  : Justin FARALAHY / MAAS — Karamako.mg / SIPEM Banque
' DATE    : 18/04/2026
' VERSION : 2.0 — Bug arivo/alina corrigé, renommé MONENLET_MG
'
' FONCTION PRINCIPALE : MONENLET_MG(Montant, [Devise])
'
' EXEMPLES VALIDÉS :
'   MONENLET_MG(1352689.60)
'   → "Sivy amby valopolo sy eninjato sy roa arivo sy dimy alina
'      sy telo hetsy sy iray tapitrisa Ariary faingo enimpolo"
'
'   MONENLET_MG(2245710096.15)
'   → "Enina amby sivifolo sy iray alina sy fito hetsy sy dimy amby
'      efapolo sy roanjato tapitrisa sy roa lavitrisa Ariary faingo
'      dimy amby folo"
'
'   MONENLET_MG(11471786)
'   → "Enina amby valopolo sy fitonjato sy arivo sy fito alina
'      sy efatra hetsy sy iraika amby folo tapitrisa Ariary"
'
' NOMENCLATURE :
'   Unités   : iray, roa, telo, efatra, dimy, enina, fito, valo, sivy
'   Dizaines : folo, roapolo, telopolo, efapolo, dimampolo,
'              enimpolo, fitopolo, valopolo, sivifolo
'   Centaines: zato, roanjato, telonjato, efajato, dimanjato,
'              eninjato, fitonjato, valonjato, sivinjato
'   Grands   : arivo (×1 000), alina (×10 000), hetsy (×100 000),
'              tapitrisa (×1 000 000), lavitrisa (×1 000 000 000)
'
' CONNECTEURS :
'   amby   — unité AVANT dizaine      ex: 9 amby valopolo = 89
'   sy     — groupe AVANT groupe sup. ex: eninjato sy roa arivo
'   faingo — virgule décimale
'
' RÈGLE SPÉCIALE arivo / alina (v2.0, BUG CORRIGÉ) :
'   arivo : multiplicateur = 1 → "arivo" seul  (jamais "iray arivo")
'           multiplicateur > 1 → "<unité> arivo"
'           ex: 1 000 → arivo  |  2 000 → roa arivo
'   alina : multiplicateur = 1 → "iray alina"  (iray obligatoire)
'           multiplicateur > 1 → "<unité> alina"
'           ex: 10 000 → iray alina  |  20 000 → roa alina
'   hetsy, tapitrisa, lavitrisa : toujours "<unité> X" même pour 1
'           ex: 100 000 → iray hetsy  |  1 M → iray tapitrisa
'
' NOTE : "iraika" est utilisé (et non "iray") avant "amby"
'        ex: 11 = iraika amby folo  |  21 = iraika amby roapolo
'
' ORDRE DE LECTURE : Ascendant (petit → grand)
'   ex: 1 352 689 = ...89... 600... 2 000... 50 000... 300 000... 1 M
'
' STRUCTURE IDENTIQUE aux modules :
'   MontantFrancaise.bas — MONENLET_PRO
'   MontantAnglaise.bas  — MONENLET_EN
'
' HISTORIQUE :
'   v1.0 — Première version (MONENLETMA) — bug : "iray arivo" au lieu de "arivo"
'   v2.0 — Renommé MONENLET_MG, règle arivo corrigée, structure unifiée
'==============================================================================


' ─────────────────────────────────────────────────────────────────────────────
' ENREGISTREMENT INFO-BULLE — À appeler une fois à l'ouverture du classeur
' ─────────────────────────────────────────────────────────────────────────────

Public Sub RegisterMONENLET_MG()
    '--------------------------------------------------------------------------
    ' Enregistre MONENLET_MG aupres d'Excel via MacroOptions.
    '
    ' IMPORTANT : Ne pas renommer en "Auto_Open".
    '   Si les trois modules (MG, FR, EN) coexistent dans le meme classeur,
    '   trois "Auto_Open" -> erreur "Ambiguous name detected".
    '   Ce Sub est appele par le module ModuleRegistration.bas.
    '--------------------------------------------------------------------------
    Dim argDesc(1) As String
    argDesc(0) = "Montant numérique à convertir (ex: 1352689.60). " & _
                 "Supporte jusqu'à 999 999 999 999 Ar (~999 lavitrisa). " & _
                 "Les valeurs négatives sont automatiquement converties en positif."
    argDesc(1) = "[Optionnel] Libellé de la devise affiché après la partie entière. " & _
                 "Défaut : ""Ariary"". Ex: ""Ar"", ""MGA"", ""Ariary""."

    Application.MacroOptions _
        Macro:="MONENLET_MG", _
        Description:="Convertit un montant numérique en lettres malgaches " & _
                      "(lecture ascendante). Connecteurs : amby (unité-dizaine), " & _
                      "sy (groupe-groupe), faingo (virgule). " & _
                      "Ex : MONENLET_MG(11471786) → ""Enina amby valopolo sy " & _
                      "fitonjato sy arivo sy fito alina sy efatra hetsy " & _
                      "sy iraika amby folo tapitrisa Ariary"".", _
        Category:="Finances MG", _
        ArgumentDescriptions:=argDesc
End Sub


' ─────────────────────────────────────────────────────────────────────────────
' FONCTION PUBLIQUE — Point d'entrée  (nom court pour usage Excel)
' ─────────────────────────────────────────────────────────────────────────────

Public Function MONENLET_MG( _
        ByVal Montant As Double, _
        Optional ByVal Devise As String = "Ariary") As String
    '--------------------------------------------------------------------------
    ' Convertit un montant numérique en lettres malgaches (lecture ascendante).
    '
    ' Syntaxe    : MONENLET_MG(Montant; [Devise])
    ' Paramètres :
    '   Montant  — valeur numérique (supporte jusqu'à ~999 lavitrisa)
    '              Les valeurs négatives sont traitées en Abs()
    '   Devise   — libellé de devise (défaut : "Ariary")
    '
    ' Retourne : chaîne en malgache, première lettre en majuscule.
    '--------------------------------------------------------------------------

    ' Protection contre les négatifs
    Montant = Abs(Montant)

    ' Cas zéro
    If Montant = 0 Then
        MONENLET_MG = "Aotra " & Devise
        Exit Function
    End If

    ' Arrondi bancaire à 2 décimales
    Montant = Round(Montant, 2)

    ' Séparation partie entière / centimes
    Dim partEntiere As Double
    Dim centimes    As Integer

    partEntiere = Int(Montant)

    ' CLng(Round()) évite la perte de précision flottante sur les centimes
    centimes = CInt(Round((Montant - partEntiere) * 100, 0))

    ' Construction du résultat
    Dim res As String
    res = NombreEntierMG(partEntiere) & " " & Devise

    If centimes > 0 Then
        res = res & " faingo " & DeuxChiffresMG(centimes)
    End If

    ' Majuscule sur la première lettre
    If Len(res) > 0 Then Mid(res, 1, 1) = UCase(Left(res, 1))

    MONENLET_MG = res
End Function


' ─────────────────────────────────────────────────────────────────────────────
' CONVERSION DE LA PARTIE ENTIÈRE (lecture ascendante)
' ─────────────────────────────────────────────────────────────────────────────

Private Function NombreEntierMG(ByVal N As Double) As String
    '--------------------------------------------------------------------------
    ' Décompose N en groupes et les assemble dans l'ordre ASCENDANT.
    '
    ' Groupes (du plus petit au plus grand) :
    '   g_ud    — unités + dizaines  (0-99)
    '   g_cent  — centaines          (1-9) → ×100
    '   g_arivo — milliers           (1-9) → ×1 000    [1 → "arivo" seul]
    '   g_alina — dizaines de mille  (1-9) → ×10 000   [1 → "iray alina"]
    '   g_hetsy — centaines de mille (1-9) → ×100 000  [1 → "iray hetsy"]
    '   millions (1-999)             → ×1 000 000      [1 → "iray tapitrisa"]
    '   milliards (1+)               → ×1 000 000 000  [1 → "iray lavitrisa"]
    '
    ' Assemblage : chaque composant est relié au suivant par " sy ".
    '--------------------------------------------------------------------------

    ' ── Extraction des groupes ──────────────────────────────────────────────
    Dim milliards As Double
    Dim millions  As Double
    Dim reste     As Double

    ' Milliards et millions via soustraction pour éviter l'imprécision
    ' de Mod avec Double au-delà de 2^31
    milliards = Int(N / 1000000000#)
    reste     = N - milliards * 1000000000#
    millions  = Int(reste / 1000000#)
    reste     = reste - millions * 1000000#

    Dim g_ud    As Integer ' 0-99
    Dim g_cent  As Integer ' 0-9
    Dim g_arivo As Integer ' 0-9
    Dim g_alina As Integer ' 0-9
    Dim g_hetsy As Integer ' 0-9

    g_ud    = CInt(reste Mod 100)
    g_cent  = CInt(Int(reste / 100) Mod 10)
    g_arivo = CInt(Int(reste / 1000) Mod 10)
    g_alina = CInt(Int(reste / 10000) Mod 10)
    g_hetsy = CInt(Int(reste / 100000) Mod 10)

    ' ── Construction de la liste ascendante ─────────────────────────────────
    Dim composants(9) As String
    Dim nComp         As Integer
    nComp = 0

    ' 1. Unités + dizaines (ex: 89 → sivy amby valopolo)
    If g_ud > 0 Then
        composants(nComp) = DeuxChiffresMG(g_ud)
        nComp = nComp + 1
    End If

    ' 2. Centaines (ex: 600 → eninjato)
    If g_cent > 0 Then
        composants(nComp) = CentaineMG(g_cent)
        nComp = nComp + 1
    End If

    ' 3. Milliers — arivo
    '    RÈGLE : g_arivo = 1 → "arivo" seul  (jamais "iray arivo")
    '            g_arivo > 1 → "<unité> arivo"
    '    ex: 1 000 → arivo  |  2 000 → roa arivo  |  9 000 → sivy arivo
    If g_arivo = 1 Then
        composants(nComp) = "arivo"
        nComp = nComp + 1
    ElseIf g_arivo > 1 Then
        composants(nComp) = UniteMG(g_arivo) & " arivo"
        nComp = nComp + 1
    End If

    ' 4. Dizaines de mille — alina
    '    RÈGLE : g_alina = 1 → "iray alina"  (iray obligatoire, contrairement à arivo)
    '            g_alina > 1 → "<unité> alina"
    '    ex: 10 000 → iray alina  |  20 000 → roa alina  |  70 000 → fito alina
    If g_alina >= 1 Then
        composants(nComp) = UniteMG(g_alina) & " alina"
        nComp = nComp + 1
    End If

    ' 5. Centaines de mille — hetsy
    '    Toujours "<unité> hetsy" même pour 1 (iray hetsy)
    '    ex: 100 000 → iray hetsy  |  400 000 → efatra hetsy
    If g_hetsy >= 1 Then
        composants(nComp) = UniteMG(g_hetsy) & " hetsy"
        nComp = nComp + 1
    End If

    ' 6. Millions — tapitrisa
    '    Le compte millions (1-999) est lui-même lu en ascendant par TroisChiffresMG.
    '    Toujours "<compte> tapitrisa" même pour 1 (iray tapitrisa)
    '    ex: 1 M → iray tapitrisa  |  11 M → iraika amby folo tapitrisa
    If millions > 0 Then
        composants(nComp) = TroisChiffresMG(CInt(millions)) & " tapitrisa"
        nComp = nComp + 1
    End If

    ' 7. Milliards — lavitrisa
    '    Toujours "<compte> lavitrisa" même pour 1 (iray lavitrisa)
    '    ex: 1 G → iray lavitrisa  |  2 G → roa lavitrisa
    If milliards > 0 Then
        composants(nComp) = TroisChiffresMG(CInt(milliards)) & " lavitrisa"
        nComp = nComp + 1
    End If

    ' ── Assemblage avec connecteur "sy" ─────────────────────────────────────
    Dim i   As Integer
    Dim res As String
    res = ""

    For i = 0 To nComp - 1
        If res = "" Then
            res = composants(i)
        Else
            res = res & " sy " & composants(i)
        End If
    Next i

    NombreEntierMG = res
End Function


' ─────────────────────────────────────────────────────────────────────────────
' CONVERSION 2 CHIFFRES (1-99)
' ─────────────────────────────────────────────────────────────────────────────

Private Function DeuxChiffresMG(ByVal N As Integer) As String
    '--------------------------------------------------------------------------
    ' Convertit un entier de 1 à 99 en malgache (ordre ascendant).
    '
    ' Règle "amby" : unité AMBY dizaine  (ex: 9 amby valopolo = 89)
    ' Forme spéciale avant amby : 1 → "iraika"  (ex: 11 = iraika amby folo)
    '                             2-9 → forme standard
    '--------------------------------------------------------------------------
    Dim u As Integer ' chiffre des unités (0-9)
    Dim d As Integer ' valeur des dizaines (0, 10, 20, ..., 90)

    u = N Mod 10
    d = N - u

    If u = 0 Then
        ' Multiple de 10 exact (ex: 80 → valopolo)
        DeuxChiffresMG = DizaineMG(d)
    ElseIf d = 0 Then
        ' Unité seule 1-9 (ex: 6 → enina)
        DeuxChiffresMG = UniteMG(u)
    Else
        ' X amby Y : unité d'abord, dizaine ensuite (ascendant)
        ' ex: 89 → sivy amby valopolo  |  11 → iraika amby folo
        DeuxChiffresMG = UniteAmbyMG(u) & " amby " & DizaineMG(d)
    End If
End Function


' ─────────────────────────────────────────────────────────────────────────────
' CONVERSION 3 CHIFFRES (1-999)
' ─────────────────────────────────────────────────────────────────────────────

Private Function TroisChiffresMG(ByVal N As Integer) As String
    '--------------------------------------------------------------------------
    ' Convertit un entier de 1 à 999 en malgache (lecture ascendante).
    ' Utilisé pour exprimer le compte de tapitrisa et lavitrisa.
    '
    ' Ordre ascendant : unités/dizaines d'abord, centaines ensuite.
    ' Reliés par "sy".
    ' ex: 245 → "dimy amby efapolo sy roanjato"
    '     11  → "iraika amby folo"
    '--------------------------------------------------------------------------
    Dim h   As Integer ' chiffre des centaines (0-9)
    Dim r   As Integer ' reste (0-99)
    Dim res As String

    h = N \ 100
    r = N Mod 100
    res = ""

    ' Unités/dizaines en premier (ascendant)
    If r > 0 Then res = DeuxChiffresMG(r)

    ' Centaines ensuite
    If h > 0 Then
        If res <> "" Then
            res = res & " sy " & CentaineMG(h)
        Else
            res = CentaineMG(h)
        End If
    End If

    TroisChiffresMG = res
End Function


' ─────────────────────────────────────────────────────────────────────────────
' HELPERS — Mots de base
' ─────────────────────────────────────────────────────────────────────────────

Private Function UniteMG(ByVal N As Integer) As String
    '--------------------------------------------------------------------------
    ' Retourne le mot malgache de l'unité N (1-9), forme de base.
    ' Utilisé seul ou comme multiplicateur de arivo/alina/hetsy/tapitrisa/lavitrisa.
    '--------------------------------------------------------------------------
    Select Case N
        Case 1: UniteMG = "iray"
        Case 2: UniteMG = "roa"
        Case 3: UniteMG = "telo"
        Case 4: UniteMG = "efatra"
        Case 5: UniteMG = "dimy"
        Case 6: UniteMG = "enina"
        Case 7: UniteMG = "fito"
        Case 8: UniteMG = "valo"
        Case 9: UniteMG = "sivy"
        Case Else: UniteMG = ""
    End Select
End Function


Private Function UniteAmbyMG(ByVal N As Integer) As String
    '--------------------------------------------------------------------------
    ' Forme de l'unité utilisée AVANT le connecteur "amby".
    ' Règle : 1 → "iraika"  (les autres unités 2-9 restent identiques)
    ' ex: 11 = iraika amby folo  |  21 = iraika amby roapolo
    '--------------------------------------------------------------------------
    If N = 1 Then
        UniteAmbyMG = "iraika"
    Else
        UniteAmbyMG = UniteMG(N)
    End If
End Function


Private Function DizaineMG(ByVal N As Integer) As String
    '--------------------------------------------------------------------------
    ' Retourne le mot malgache de la dizaine N (10, 20, ..., 90).
    '--------------------------------------------------------------------------
    Select Case N
        Case 10: DizaineMG = "folo"
        Case 20: DizaineMG = "roapolo"
        Case 30: DizaineMG = "telopolo"
        Case 40: DizaineMG = "efapolo"
        Case 50: DizaineMG = "dimampolo"
        Case 60: DizaineMG = "enimpolo"
        Case 70: DizaineMG = "fitopolo"
        Case 80: DizaineMG = "valopolo"
        Case 90: DizaineMG = "sivifolo"
        Case Else: DizaineMG = ""
    End Select
End Function


Private Function CentaineMG(ByVal N As Integer) As String
    '--------------------------------------------------------------------------
    ' Retourne le mot malgache de la centaine N×100 (N de 1 à 9).
    '--------------------------------------------------------------------------
    Select Case N
        Case 1: CentaineMG = "zato"
        Case 2: CentaineMG = "roanjato"
        Case 3: CentaineMG = "telonjato"
        Case 4: CentaineMG = "efajato"
        Case 5: CentaineMG = "dimanjato"
        Case 6: CentaineMG = "eninjato"
        Case 7: CentaineMG = "fitonjato"
        Case 8: CentaineMG = "valonjato"
        Case 9: CentaineMG = "sivinjato"
        Case Else: CentaineMG = ""
    End Select
End Function


' ─────────────────────────────────────────────────────────────────────────────
' PROCÉDURE DE TEST (à exécuter dans l'IDE VBA : F5 sur Sub TestMontantMG)
' ─────────────────────────────────────────────────────────────────────────────

Public Sub TestMontantMG()
    '--------------------------------------------------------------------------
    ' Lance une série de tests et affiche les résultats dans la fenêtre
    ' Exécution (Ctrl+G dans l'IDE VBA).
    '--------------------------------------------------------------------------
    Dim sep As String
    sep = String(72, "-")

    Debug.Print sep
    Debug.Print "TEST — MONENLET_MG v2.0"
    Debug.Print sep

    ' ── Exemples de référence ───────────────────────────────────────────────
    Debug.Print ""
    Debug.Print "[RÉFÉRENCES]"

    Debug.Print "[REF 1]  1 352 689,60 Ar"
    Debug.Print "Attendu : Sivy amby valopolo sy eninjato sy roa arivo sy dimy alina sy telo hetsy sy iray tapitrisa Ariary faingo enimpolo"
    Debug.Print "Obtenu  : " & MONENLET_MG(1352689.6)
    Debug.Print ""

    Debug.Print "[REF 2]  2 245 710 096,15 Ar"
    Debug.Print "Attendu : Enina amby sivifolo sy iray alina sy fito hetsy sy dimy amby efapolo sy roanjato tapitrisa sy roa lavitrisa Ariary faingo dimy amby folo"
    Debug.Print "Obtenu  : " & MONENLET_MG(2245710096.15)
    Debug.Print ""

    Debug.Print "[REF 3]  11 471 786,00 Ar  (bug v1.0 corrigé)"
    Debug.Print "Attendu : Enina amby valopolo sy fitonjato sy arivo sy fito alina sy efatra hetsy sy iraika amby folo tapitrisa Ariary"
    Debug.Print "Obtenu  : " & MONENLET_MG(11471786)
    Debug.Print ""

    ' ── Règle arivo / alina ─────────────────────────────────────────────────
    Debug.Print sep
    Debug.Print "[RÈGLE arivo / alina]"
    Debug.Print "1 000        → " & MONENLET_MG(1000, "Ariary")     ' arivo (sans iray)
    Debug.Print "2 000        → " & MONENLET_MG(2000, "Ariary")     ' roa arivo
    Debug.Print "9 000        → " & MONENLET_MG(9000, "Ariary")     ' sivy arivo
    Debug.Print "10 000       → " & MONENLET_MG(10000, "Ariary")    ' iray alina (avec iray)
    Debug.Print "20 000       → " & MONENLET_MG(20000, "Ariary")    ' roa alina
    Debug.Print "11 000       → " & MONENLET_MG(11000, "Ariary")    ' arivo sy iray alina
    Debug.Print "100 000      → " & MONENLET_MG(100000, "Ariary")   ' iray hetsy
    Debug.Print "1 000 000    → " & MONENLET_MG(1000000, "Ariary")  ' iray tapitrisa
    Debug.Print "1 000 000 000 → " & MONENLET_MG(1000000000, "Ariary") ' iray lavitrisa
    Debug.Print ""

    ' ── Cas généraux ────────────────────────────────────────────────────────
    Debug.Print sep
    Debug.Print "[CAS GÉNÉRAUX]"
    Debug.Print "0          → " & MONENLET_MG(0)
    Debug.Print "1          → " & MONENLET_MG(1)
    Debug.Print "11         → " & MONENLET_MG(11)
    Debug.Print "21         → " & MONENLET_MG(21)
    Debug.Print "100        → " & MONENLET_MG(100)
    Debug.Print "500,75     → " & MONENLET_MG(500.75)
    Debug.Print "1 000,01   → " & MONENLET_MG(1000.01)
    Debug.Print sep
End Sub
