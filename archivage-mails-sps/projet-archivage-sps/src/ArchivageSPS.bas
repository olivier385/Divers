Attribute VB_Name = "ArchivageSPS"
'==============================================================================
'  ARCHIVAGE DES MAILS ET PIECES JOINTES DE LA BOITE secretariat.sps
'  Groupe DEGOUY - Pôle SPS - version 1.10 du 25/09/2026
'
'  Nouveauté de la version 1.10 :
'   - CORRECTION : les mails du premier jour de la période reçus avant midi
'     étaient ignorés (filtre de date en « h:nn AMPM » mal lu sur un Windows
'     français). Filtre désormais en 24 h, avec contrôle exact mail par mail.
'
'  Nouveautés de la version 1.9 :
'   - aucun fichier n'est jamais écrasé : existence vérifiée juste avant
'     chaque écriture (mails, pièces jointes, journal, diagnostic) ;
'   - journal et diagnostic nommés avec l'utilisateur Windows et l'heure à
'     la seconde : chaque poste garde le sien ;
'   - la simulation mémorise les fichiers « à copier » : un fichier présent
'     dans plusieurs mails de la période n'est compté qu'une fois ;
'   - un mail déjà archivé en .eml (script pst_archive.py) n'est pas repris ;
'   - images intégrées (Content-ID) et images image* / logo* ignorées ;
'   - photos (.heic, .jpg...) d'un mail qui cite plusieurs affaires :
'     signalées dans le journal, non copiées (COPIER_PHOTOS_MULTI_AFFAIRES) ;
'   - correspondances cherchées en mots entiers ; apostrophe ’ et Œ reconnus ;
'   - liste des mots-clés du nom de fichier complétée (ICP, OUVERTURE,
'     CSPS, PV DE REUNION...).
'
'  Rôle :
'   - parcourt TOUS les dossiers de courrier de la boîte secretariat.sps
'     (réception, envoyés, dossiers de classement par coordonnateur...),
'     y compris Éléments supprimés et Boîte d'envoi (sauf courrier indésirable
'     et brouillons),
'     entre DATE_DEBUT et DATE_FIN ;
'   - repère le n° d'affaire 7.AAAA.NNN dans l'objet, sinon dans la table
'     CORRESPONDANCES (mots-clés -> affaire), sinon dans le début du texte ;
'   - enregistre le mail en .msg dans U:\S.P.S. AA\<affaire>\06_MAILS\Reçus
'     ou \Envoyés ;
'   - enregistre chaque pièce jointe dans le bon dossier de l'affaire :
'       RJ / VI / CR SPS           -> 01_RJ et IC\RJ (sous-dossier de l'année s'il existe)
'       IC / VIC / inspection      -> 01_RJ et IC\IC (sous-dossier de l'année s'il existe)
'       PPSPS / PIC                -> 02_PPSPS et PIC
'       CR de la maîtrise d'œuvre  -> 03_CR CHANTIER
'       CISSCT                     -> 04_CISSCT
'       DIUO                       -> 07_DIUO
'       PGC                        -> 09_PGC
'       APS-AVP / APD / DCE        -> 10 / 11 / 13
'       autres                     -> 05 (DOCS TECH ou DIVERS)
'   - ne remplace JAMAIS un fichier et évite les doublons : un fichier est
'     ignoré s'il existe déjà sous le même nom, ou sous un nom proche (date
'     en tête, « (1) »...) avec la même taille, dans le dossier concerné ;
'   - une pièce jointe dont le nom cite une autre affaire est rangée dans
'     cette autre affaire ;
'   - écrit un journal .csv (ouvrable dans Excel) dans JOURNAL_DOSSIER.
'
'  Mode SIMULATION = True : rien n'est copié, seul le journal est produit.
'  Passer SIMULATION à False une fois le journal validé.
'
'  Installation (Outlook classique pour Windows ; le « nouvel Outlook » ne
'  gère pas les macros) :
'   1. Outlook > Alt+F11 > Fichier > Importer un fichier > ArchivageSPS.bas
'   2. Outils > Références : rien à cocher (liaisons tardives).
'   3. Si les macros sont bloquées : Fichier > Options > Centre de gestion
'      de la confidentialité > Paramètres des macros > « Notifications pour
'      toutes les macros », puis redémarrer Outlook.
'   4. Lancer la macro « LancerArchivage » (Alt+F8 dans Outlook).
'==============================================================================
Option Explicit

'---------------------------- PARAMETRES --------------------------------------
Private Const BOITE As String = "secretariat.sps@degouy.fr"
Private Const BOITE_NOM_AFFICHE As String = "Secretariat-SPS"   ' autre nom possible de la boîte dans Outlook
Private Const RACINE As String = "U:\"                           ' contient les dossiers « S.P.S. 14 » à « S.P.S. 26 »
Private Const DATE_DEBUT As String = "2026-09-21"                ' incluse (AAAA-MM-JJ)
Private Const DATE_FIN As String = "2026-09-26"                   ' exclue  (AAAA-MM-JJ)
Private Const SIMULATION As Boolean = True                       ' True = aucun fichier copié
Private Const INCLURE_SOUS_DOSSIERS As Boolean = True
Private Const TRAITER_ENVOYES As Boolean = True
Private Const JOURNAL_DOSSIER As String = "U:\S.P.S. 26\_ARCHIVAGE MAILS"
Private Const MAX_CHEMIN As Long = 250
Private Const VERSION As String = "1.10"
' Photos jointes à un mail qui cite plusieurs affaires : False = signalées dans le journal, non copiées
Private Const COPIER_PHOTOS_MULTI_AFFAIRES As Boolean = False
' Mots-clés (objet ou adresse de l'expéditeur) -> n° d'affaire, pour les mails sans n° d'affaire.
' Format : "mot-clé=7.AAAA.NNN;mot-clé=7.AAAA.NNN" (majuscules/minuscules et ponctuation indifférentes,
' recherche en mots entiers). À compléter.
Private Const CORRESPONDANCES As String = _
    "54 RUE DE ROMAINVILLE=7.2025.041;STEP CHAUMES=7.2023.463;PONT D'IVRY=7.2025.397;" & _
    "CHUGPN=7.2022.198;BONDOUFLE=7.2021.055;PONT AMAR=7.2021.455;hbarchitectes.fr=7.2021.455;" & _
    "BEAUDELAIRE=7.2026.019;BAUDELAIRE=7.2026.019;BILLETTES=7.2025.042;COLLEGE HONORE DE BALZAC=7.2023.002;" & _
    "FOYER PARIS DUMAS=7.2025.345;BLOMET=7.2019.347;atelierboteko=7.2019.347;GLACIERE=7.2022.500;" & _
    "MERCOEUR=7.2023.218;wao.paris=7.2023.218;STADE NAUTIQUE=7.2026.272;BD NEY=7.2021.105;" & _
    "SAINT-BERNARD-DE-LA-CHAPELLE=7.2023.282;EPHE=7.2024.324;RESTO DU COEUR=7.2025.386;RESTOS DU COEUR=7.2025.386;" & _
    "LEG SARTROUVILLE=7.2020.533;STATION FOCH=7.2024.384;CLICHY-SOUS-BOIS=7.2022.448;" & _
    "HOTEL DIEU=7.2018.194;LOURCINE=7.2023.437;RUE BERTHIER=7.2020.512"
'------------------------------------------------------------------------------

Private fso As Object
Private rxAffaire As Object
Private cacheAffaires As Object      ' n° affaire -> chemin du dossier
Private cacheIndex As Object         ' chemin de dossier -> dictionnaire des noms de fichiers
Private journal As Collection
Private dDebut As Date, dFin As Date
Private nbMails As Long, nbMailsCopies As Long, nbPJ As Long, nbPJCopiees As Long
Private nbDeja As Long, nbSansAffaire As Long, nbErreurs As Long, nbDossiers As Long, nbPhotos As Long

'==============================================================================
Public Sub LancerArchivage()
    Dim st As Object, fIn As Object, fSent As Object, t As Single

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set cacheAffaires = CreateObject("Scripting.Dictionary")
    Set cacheIndex = CreateObject("Scripting.Dictionary")
    Set journal = New Collection
    Set rxAffaire = CreateObject("VBScript.RegExp")
    rxAffaire.Global = True
    rxAffaire.IgnoreCase = True
    rxAffaire.Pattern = "(^|[^0-9])7[ ._-]{0,3}(20[0-9]{2})[ ._-]{1,3}([0-9]{3})(?![0-9])"

    dDebut = DateSerial(CInt(Left(DATE_DEBUT, 4)), CInt(Mid(DATE_DEBUT, 6, 2)), CInt(Right(DATE_DEBUT, 2)))
    dFin = DateSerial(CInt(Left(DATE_FIN, 4)), CInt(Mid(DATE_FIN, 6, 2)), CInt(Right(DATE_FIN, 2)))
    nbMails = 0: nbMailsCopies = 0: nbPJ = 0: nbPJCopiees = 0: nbDeja = 0: nbSansAffaire = 0: nbErreurs = 0: nbDossiers = 0: nbPhotos = 0

    journal.Add "Date;Sens;Expéditeur;Objet;Affaire;Dossier affaire;Élément;Fichier;Destination;Statut;Dossier Outlook"
    journal.Add ";;;;;;;;;" & Csv("VERSION " & VERSION & " - utilisateur " & Environ("USERNAME") & " - poste " & Environ("COMPUTERNAME") & _
                IIf(SIMULATION, " - SIMULATION", " - COPIE RÉELLE")) & ";"

    If Dir(RACINE & "S.P.S. 26", vbDirectory) = "" Then
        MsgBox "Le lecteur " & RACINE & " n'est pas accessible (dossier « S.P.S. 26 » introuvable).", vbCritical
        Exit Sub
    End If

    t = Timer
    Set st = TrouverBoite()
    If Not st Is Nothing Then
        Set fIn = st.GetDefaultFolder(6)            ' olFolderInbox
        If TRAITER_ENVOYES Then
            On Error Resume Next
            Set fSent = st.GetDefaultFolder(5)      ' olFolderSentMail
            On Error GoTo 0
        End If
    Else
        ' Boîte non ouverte comme compte : accès partagé direct
        On Error Resume Next
        Dim rcp As Object
        Set rcp = Application.Session.CreateRecipient(BOITE)
        rcp.Resolve
        Set fIn = Application.Session.GetSharedDefaultFolder(rcp, 6)
        If TRAITER_ENVOYES Then Set fSent = Application.Session.GetSharedDefaultFolder(rcp, 5)
        On Error GoTo 0
    End If

    If fIn Is Nothing Then
        MsgBox "Boîte " & BOITE & " introuvable dans Outlook." & vbCrLf & _
               "Vérifiez qu'elle apparaît dans le volet des dossiers.", vbCritical
        Exit Sub
    End If

    If Not st Is Nothing Then
        TraiterArborescence st.GetRootFolder, st
    Else
        TraiterDossier fIn, "Reçus"
        If Not fSent Is Nothing Then TraiterDossier fSent, "Envoyés"
    End If

    Dim fichierJournal As String
    fichierJournal = EcrireJournal()

    MsgBox "Version " & VERSION & " - " & IIf(SIMULATION, "SIMULATION terminée (aucun fichier copié)", "Archivage terminé") & vbCrLf & vbCrLf & _
           "Dossiers Outlook parcourus : " & nbDossiers & vbCrLf & _
           "Période : " & Format(dDebut, "dd/mm/yyyy") & " au " & Format(dFin - 1, "dd/mm/yyyy") & vbCrLf & _
           "Mails traités : " & nbMails & vbCrLf & _
           "Mails " & IIf(SIMULATION, "à enregistrer", "enregistrés") & " : " & nbMailsCopies & vbCrLf & _
           "Pièces jointes examinées : " & nbPJ & vbCrLf & _
           "Pièces jointes " & IIf(SIMULATION, "à copier", "copiées") & " : " & nbPJCopiees & vbCrLf & _
           "Déjà présents sur le serveur : " & nbDeja & vbCrLf & _
           "Mails sans n° d'affaire : " & nbSansAffaire & vbCrLf & _
           "Photos de mails multi-affaires (à classer à la main) : " & nbPhotos & vbCrLf & _
           "Erreurs : " & nbErreurs & vbCrLf & vbCrLf & _
           "Durée : " & Format((Timer - t) / 60, "0.0") & " min" & vbCrLf & _
           "Journal : " & fichierJournal, vbInformation, "Archivage SPS"
End Sub

'==============================================================================
Private Function TrouverBoite() As Object
    Dim st As Object, nom As String
    For Each st In Application.Session.Stores
        nom = LCase(st.DisplayName)
        If InStr(nom, LCase(BOITE)) > 0 Or InStr(nom, LCase(BOITE_NOM_AFFICHE)) > 0 Then
            Set TrouverBoite = st
            Exit Function
        End If
    Next
End Function

' Parcourt tous les dossiers de courrier de la boîte (dossiers de classement compris).
' 1) on dresse d'abord la liste complète des dossiers, 2) puis on traite chaque dossier.
Private Sub TraiterArborescence(ByVal racineBoite As Object, ByVal st As Object)
    Dim exclus As Object, idEnvoyes As String, v As Variant
    Dim liste As New Collection, i As Long, st2 As Object
    Set exclus = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    For Each v In Array(23, 16, 9, 10, 11, 12, 13)   ' indésirables, brouillons, calendrier, contacts, journal, notes, tâches
        exclus(st.GetDefaultFolder(v).EntryID) = True
    Next
    idEnvoyes = st.GetDefaultFolder(5).EntryID
    On Error GoTo 0
    ' même méthode que le diagnostic (qui, lui, voit tous les sous-dossiers) :
    ' parcours des boîtes via Session.Stores, ouverture de chaque dossier puis énumération For Each
    For Each st2 In Application.Session.Stores
        If st2.StoreID = st.StoreID Then
            CollecteDiag st2.GetRootFolder, liste
            Exit For
        End If
    Next
    journal.Add ";;;;;;;;;" & liste.Count & " DOSSIERS TROUVÉS;"
    Dim f As Object, sens As String, nom As String
    For i = 1 To liste.Count
        Set f = liste(i)
        On Error Resume Next
        nom = LCase(f.Name)
        If exclus.Exists(f.EntryID) Then GoTo Suivant
        If f.DefaultItemType <> 0 Then GoTo Suivant
        If nom Like "*problèmes de synchronisation*" Or nom Like "*sync issues*" Or nom Like "*historique des conversations*" _
           Or nom Like "*conversation history*" Or nom Like "*flux rss*" Or nom Like "*rss feeds*" Or nom Like "*yammer*" _
           Or nom Like "*conflits*" Or nom Like "*défaillances*" Then GoTo Suivant
        If f.FolderPath = racineBoite.FolderPath Then GoTo Suivant
        sens = "Reçus"
        If InStr(f.FolderPath, "\" & st.GetDefaultFolder(5).Name) > 0 Then sens = "Envoyés"
        On Error GoTo 0
        TraiterDossier f, sens, False
Suivant:
        On Error GoTo 0
    Next
End Sub

' Copie conforme du parcours du diagnostic
Private Sub CollecteDiag(ByVal f As Object, ByVal liste As Collection)
    Dim sf As Object, n As Long, total As Long, filtre As String
    On Error Resume Next
    total = f.Items.Count
    filtre = FiltrePeriode(dDebut, dFin)
    n = f.Items.Restrict(filtre).Count
    liste.Add f
    journal.Add ";;;;;;;;;DOSSIER VU (" & total & " éléments, " & n & " sur la période);" & Csv(f.FolderPath)
    For Each sf In f.Folders
        CollecteDiag sf, liste
    Next
End Sub

Private Sub CollecterDossiers(ByVal f As Object, ByVal exclus As Object, ByVal idEnvoyes As String, _
                             ByVal sens As String, ByVal liste As Collection)
    Dim sf As Object, nom As String, id As String, typ As Long, nb As Long, k As Long, chemin As String, errTxt As String
    Dim vus As Object
    On Error Resume Next
    id = f.EntryID
    chemin = f.FolderPath
    typ = -1: typ = f.DefaultItemType
    nom = LCase(f.Name)
    ' Outlook ne charge les sous-dossiers d'une boîte partagée qu'après accès au contenu du dossier :
    ' on « ouvre » donc le dossier (lecture du nombre d'éléments) avant de lister ses sous-dossiers
    Dim nbEl As Long: nbEl = f.Items.Count
    DoEvents
    nb = -1: nb = f.Folders.Count
    If Err.Number <> 0 Then errTxt = " ERREUR : " & Err.Description: Err.Clear
    journal.Add ";;;;;;;;;DOSSIER VU (type " & typ & ", " & nb & " sous-dossier(s))" & errTxt & ";" & Csv(chemin)
    If exclus.Exists(id) Then journal.Add ";;;;;;;;;DOSSIER EXCLU;" & Csv(chemin): Exit Sub
    If typ <> 0 Then Exit Sub                                   ' dossiers non courrier
    If nom Like "*problèmes de synchronisation*" Or nom Like "*sync issues*" Or nom Like "*historique des conversations*" _
       Or nom Like "*conversation history*" Or nom Like "*flux rss*" Or nom Like "*rss feeds*" Or nom Like "*yammer*" Then Exit Sub
    If id = idEnvoyes Then sens = "Envoyés"
    liste.Add Array(f, sens)
    ' sous-dossiers : deux méthodes d'énumération, sans doublon
    Set vus = CreateObject("Scripting.Dictionary")
    For k = 1 To nb
        Set sf = Nothing
        Set sf = f.Folders.Item(k)
        If Not sf Is Nothing Then
            If Not vus.Exists(sf.EntryID) Then vus(sf.EntryID) = True: CollecterDossiers sf, exclus, idEnvoyes, sens, liste
        End If
    Next
    For Each sf In f.Folders
        If Not vus.Exists(sf.EntryID) Then vus(sf.EntryID) = True: CollecterDossiers sf, exclus, idEnvoyes, sens, liste
    Next
End Sub

' Filtre Outlook sur la date de réception, en heure 24 h.
' (v1.9 et avant : « h:nn AMPM » ; sur un Windows français, sans indicateur AM/PM, minuit s'écrivait « 12:00 »
'  et Outlook le lisait comme midi : tous les mails du premier jour avant 12 h étaient ignorés.)
Private Function FiltrePeriode(ByVal d1 As Date, ByVal d2 As Date) As String
    FiltrePeriode = "[ReceivedTime] >= '" & Format(d1, "ddddd hh:nn") & "' AND [ReceivedTime] < '" & Format(d2, "ddddd hh:nn") & "'"
End Function

Private Sub TraiterDossier(ByVal dossier As Object, ByVal sens As String, Optional ByVal recursif As Boolean = True)
    Dim elements As Object, it As Object, i As Long, filtre As String, sf As Object

    ' fenêtre élargie d'un jour de chaque côté : la date exacte est contrôlée plus bas, mail par mail
    filtre = FiltrePeriode(dDebut - 1, dFin + 1)
    On Error Resume Next
    Set elements = dossier.Items.Restrict(filtre)
    If Err.Number <> 0 Then Err.Clear: Set elements = dossier.Items
    On Error GoTo 0

    nbDossiers = nbDossiers + 1
    journal.Add ";;;;;;;;;DOSSIER PARCOURU (" & elements.Count & " élément(s) entre la veille et le lendemain de la période);" & Csv(dossier.FolderPath)
    Debug.Print "Dossier " & dossier.FolderPath & " : " & elements.Count & " élément(s)"
    For i = 1 To elements.Count
        Set it = elements(i)
        If it.Class = 43 Then                    ' olMail
            If it.ReceivedTime >= dDebut And it.ReceivedTime < dFin Then TraiterMail it, sens
        End If
        If i Mod 10 = 0 Then DoEvents
    Next

    If INCLURE_SOUS_DOSSIERS And recursif Then
        For Each sf In dossier.Folders
            TraiterDossier sf, sens
        Next
    End If
End Sub

'==============================================================================
Private Sub TraiterMail(ByVal m As Object, ByVal sens As String)
    Dim objet As String, expediteur As String, externe As Boolean
    Dim affaires As Collection, a As Variant, chemin As String
    Dim att As Object, nomPJ As String, cible As String, element As String
    Dim nPJ As Collection, dest As String, nomMail As String, cheminPJ As String, affPJ As String
    Dim ext As String, taille As Long

    On Error GoTo Erreur
    objet = Nz(m.Subject)
    expediteur = AdresseExpediteur(m)
    If MailIgnore(objet, expediteur) Then Exit Sub
    nbMails = nbMails + 1
    ' un mail émis par la boîte secretariat est un mail envoyé, quel que soit le dossier où il se trouve
    If LCase(expediteur) = LCase(BOITE) Then sens = "Envoyés"
    externe = (InStr(LCase(expediteur), "@degouy.fr") = 0)
    ' un mail transféré (TR / FW) porte en général un document reçu de l'extérieur
    If Not externe Then
        If LCase(Left(Trim(objet), 3)) = "tr:" Or LCase(Left(Trim(objet), 3)) = "fw:" Or LCase(Left(Trim(objet), 4)) = "fwd:" Then externe = True
    End If

    Set affaires = ExtraireAffaires(objet)
    If affaires.Count = 0 Then Set affaires = AffaireParCorrespondance(objet & " " & expediteur)
    If affaires.Count = 0 Then Set affaires = ExtraireAffaires(Left(Nz(m.Body), 3000), True)
    If affaires.Count = 0 Then
        nbSansAffaire = nbSansAffaire + 1
        Journaliser m, sens, expediteur, "", "", "Mail", "", "", "SANS N° AFFAIRE - à classer à la main"
        Exit Sub
    End If

    For Each a In affaires
        chemin = DossierAffaire(CStr(a))
        If chemin = "" Then
            Journaliser m, sens, expediteur, CStr(a), "", "Mail", "", "", "DOSSIER AFFAIRE INTROUVABLE"
            GoTo Suivante
        End If

        '--- 1. le mail lui-même -> 06_MAILS\Reçus ou \Envoyés
        dest = SousDossier(chemin, "06", "06_MAILS") & "\" & sens
        nomMail = Format(m.ReceivedTime, "yyyy-mm-dd hhnn") & " - " & NettoyerNom(objet)
        nomMail = Tronquer(dest, nomMail, ".msg")
        If MailDejaArchive(SousDossier(chemin, "06", "06_MAILS"), objet, nomMail) Or fso.FileExists(dest & "\" & nomMail) Then
            nbDeja = nbDeja + 1
            Journaliser m, sens, expediteur, CStr(a), chemin, "Mail", nomMail, dest, "DÉJÀ PRÉSENT"
        Else
            If Not SIMULATION Then
                CreerDossier dest
                m.SaveAs dest & "\" & nomMail, 9        ' olMSGUnicode
            End If
            ' mémorisé aussi en simulation : un même nom n'est compté qu'une fois
            AjouterIndex SousDossier(chemin, "06", "06_MAILS"), nomMail, 0
            nbMailsCopies = nbMailsCopies + 1
            Journaliser m, sens, expediteur, CStr(a), chemin, "Mail", nomMail, dest, IIf(SIMULATION, "À ENREGISTRER", "ENREGISTRÉ")
        End If

        '--- 2. les pièces jointes
        For Each att In m.Attachments
            If Not PJIgnoree(att) Then
                nomPJ = NettoyerNom(att.FileName)
                ' une PJ dont le nom cite une autre affaire est rangée dans cette affaire
                cheminPJ = chemin: affPJ = CStr(a)
                Set nPJ = ExtraireAffaires(att.FileName)
                If nPJ.Count > 0 Then
                    If Not Contient(nPJ, CStr(a)) Then
                        If Contient(affaires, CStr(nPJ(1))) Then GoTo PJSuivante   ' traitée avec son affaire
                        If DossierAffaire(CStr(nPJ(1))) <> "" Then cheminPJ = DossierAffaire(CStr(nPJ(1))): affPJ = CStr(nPJ(1))
                    End If
                End If
                ' photo d'un mail qui cite plusieurs affaires, sans affaire dans son nom : on ne sait pas où la ranger
                If affaires.Count > 1 And nPJ.Count = 0 And EstPhoto(nomPJ) And Not COPIER_PHOTOS_MULTI_AFFAIRES Then
                    If a = affaires(1) Then                  ' une seule ligne de journal par photo
                        nbPhotos = nbPhotos + 1
                        Journaliser m, sens, expediteur, JoindreAffaires(affaires), "", "Photo", nomPJ, "", _
                                    "PHOTO - mail multi-affaires, à classer à la main"
                    End If
                    GoTo PJSuivante
                End If
                nbPJ = nbPJ + 1
                cible = DestinationPJ(cheminPJ, nomPJ, objet, externe, m.ReceivedTime, element)
                ext = fso.GetExtensionName(nomPJ)
                If ext <> "" Then ext = "." & ext
                nomPJ = Tronquer(cible, fso.GetBaseName(nomPJ), ext)
                If FichierDejaPresent(RacineCategorie(cheminPJ, cible), nomPJ, att) Or fso.FileExists(cible & "\" & nomPJ) Then
                    nbDeja = nbDeja + 1
                    Journaliser m, sens, expediteur, affPJ, cheminPJ, element, nomPJ, cible, "DÉJÀ PRÉSENT"
                Else
                    If Not SIMULATION Then
                        CreerDossier cible
                        att.SaveAsFile cible & "\" & nomPJ
                        taille = FileLen(cible & "\" & nomPJ)
                    Else
                        taille = TailleExacte(att, nomPJ)
                    End If
                    ' mémorisé aussi en simulation : un même fichier n'est compté qu'une fois
                    AjouterIndex RacineCategorie(cheminPJ, cible), nomPJ, taille
                    nbPJCopiees = nbPJCopiees + 1
                    Journaliser m, sens, expediteur, affPJ, cheminPJ, element, nomPJ, cible, IIf(SIMULATION, "À COPIER", "COPIÉ")
                End If
            End If
PJSuivante:
        Next att
Suivante:
    Next a
    Exit Sub

Erreur:
    nbErreurs = nbErreurs + 1
    On Error Resume Next
    Journaliser m, sens, expediteur, "", "", "", "", "", "ERREUR : " & Err.Description
End Sub

'==============================================================================
'  Classement des pièces jointes
'==============================================================================
Private Function DestinationPJ(ByVal chemin As String, ByVal nomPJ As String, ByVal objet As String, _
                               ByVal externe As Boolean, ByVal dateMail As Date, ByRef element As String) As String
    Dim u As String, base As String, d As String

    u = " " & Normaliser(nomPJ) & " "
    ' si le nom du fichier ne dit rien, on s'appuie sur l'objet du mail
    If Not ContientMotCle(u) Then u = u & " " & Normaliser(objet) & " "

    If MotPresent(u, "PPSPS|PIC|PLAN D INSTALLATION DE CHANTIER") Then
        element = "PPSPS / PIC": DestinationPJ = SousDossier(chemin, "02", "02_PPSPS et PIC")
    ElseIf MotPresent(u, "AVIS") And MotPresent(u, "PRO") Then
        element = "Avis PRO": DestinationPJ = SousDossier(chemin, "12", "12_PRO")
    ElseIf MotPresent(u, "AVIS") And MotPresent(u, "DCE") Then
        element = "Avis DCE": DestinationPJ = SousDossier(chemin, "13", "13_DCE")
    ElseIf MotPresent(u, "AVIS") And MotPresent(u, "APD") Then
        element = "Avis APD": DestinationPJ = SousDossier(chemin, "11", "11_APD")
    ElseIf MotPresent(u, "AVIS") And MotPresent(u, "APS|AVP") Then
        element = "Avis APS / AVP": DestinationPJ = SousDossier(chemin, "10", "10_APS et AVP")
    ElseIf MotPresent(u, "DIUO") Then
        element = "DIUO": DestinationPJ = SousDossier(chemin, "07", "07_DIUO")
    ElseIf MotPresent(u, "CISSCT") Then
        element = "CISSCT": DestinationPJ = SousDossier(chemin, "04", "04_CISSCT")
    ElseIf MotPresent(u, "PGC") Then
        element = "PGC": DestinationPJ = SousDossier(chemin, "09", "09_PGC")
    ElseIf MotPresent(u, "LIVRET|LIVRET D ACCUEIL|LIVRET D ACCEUIL") Then
        element = "Autre document"
        DestinationPJ = SousDossier(chemin, "05", "")
        If DestinationPJ = "" Then DestinationPJ = SousDossier(chemin, "05", "05_DIVERS")
    ElseIf MotPresent(u, "RJ|REGISTRE") Then
        element = "RJ"
        base = SousDossier(chemin, "01", "01_RJ et IC")
        DestinationPJ = SousDossierAnnee(SousDossierNomme(base, "RJ"), dateMail)
    ElseIf MotPresent(u, "IC|ICMOD|VIC|ICP|INSPECTION COMMUNE|INSPECTION") Then
        element = "IC"
        base = SousDossier(chemin, "01", "01_RJ et IC")
        DestinationPJ = SousDossierAnnee(SousDossierNomme(base, "IC"), dateMail)
    ElseIf externe And MotPresent(u, "CR[0-9]*|COMPTE RENDU|COMPTERENDU|PV DE REUNION") And Not MotPresent(u, "SPS|CSPS") Then
        element = "CR maîtrise d'œuvre": DestinationPJ = SousDossier(chemin, "03", "03_CR CHANTIER")
    ElseIf MotPresent(u, "VI|VI[0-9]+|VISITE|CR SPS|CSPS|OUVERTURE|CR[0-9]*") Then
        element = "RJ"
        base = SousDossier(chemin, "01", "01_RJ et IC")
        DestinationPJ = SousDossierAnnee(SousDossierNomme(base, "RJ"), dateMail)
    ElseIf MotPresent(u, "DCE") Then
        element = "DCE": DestinationPJ = SousDossier(chemin, "13", "13_DCE")
    ElseIf MotPresent(u, "APD") Then
        element = "APD": DestinationPJ = SousDossier(chemin, "11", "11_APD")
    ElseIf MotPresent(u, "APS|AVP") Then
        element = "APS / AVP": DestinationPJ = SousDossier(chemin, "10", "10_APS et AVP")
    Else
        element = "Autre document"
        d = SousDossier(chemin, "05", "")
        If d = "" Then d = SousDossier(chemin, "05", "05_DIVERS")
        DestinationPJ = d
    End If
End Function

Private Function ContientMotCle(ByVal u As String) As Boolean
    ContientMotCle = MotPresent(u, "PPSPS|PIC|PLAN D INSTALLATION DE CHANTIER|DIUO|CISSCT|PGC|IC|ICMOD|VIC|ICP|INSPECTION|RJ|VI|VI[0-9]+|REGISTRE|VISITE|OUVERTURE|CSPS|CR SPS|CR[0-9]*|COMPTE RENDU|COMPTERENDU|PV DE REUNION|DCE|APD|APS|AVP|LIVRET|AVIS")
End Function

Private Function MotPresent(ByVal texte As String, ByVal mots As String) As Boolean
    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = True
    rx.Pattern = "(^|[^A-Z0-9])(" & mots & ")([^A-Z0-9]|$)"
    MotPresent = rx.Test(texte)
End Function

' Racine de catégorie (ex. 01_RJ et IC) pour la recherche de doublons
Private Function RacineCategorie(ByVal chemin As String, ByVal cible As String) As String
    Dim reste As String, p As Long
    reste = Mid(cible, Len(chemin) + 2)
    p = InStr(reste, "\")
    If p > 0 Then reste = Left(reste, p - 1)
    RacineCategorie = chemin & "\" & reste
End Function

'==============================================================================
'  Dossiers du serveur
'==============================================================================
Private Function DossierAffaire(ByVal num As String) As String
    Dim annee As String, base As String, d As String, suite As String
    If cacheAffaires.Exists(num) Then DossierAffaire = cacheAffaires(num): Exit Function
    annee = Mid(num, 3, 4)
    base = RACINE & "S.P.S. " & Right(annee, 2)
    If Dir(base, vbDirectory) <> "" Then
        d = Dir(base & "\" & num & "*", vbDirectory)
        Do While d <> ""
            suite = Mid(d, Len(num) + 1, 1)
            If suite = "" Or Not (suite Like "[0-9]") Then
                If (GetAttr(base & "\" & d) And vbDirectory) = vbDirectory Then
                    DossierAffaire = base & "\" & d
                    Exit Do
                End If
            End If
            d = Dir()
        Loop
    End If
    cacheAffaires(num) = DossierAffaire
End Function

' Sous-dossier dont le nom commence par « prefixe » ; s'il n'existe pas, on renvoie « defaut » (créé plus tard si besoin)
Private Function SousDossier(ByVal chemin As String, ByVal prefixe As String, ByVal defaut As String) As String
    Dim f As Object
    For Each f In fso.GetFolder(chemin).SubFolders
        If Left(f.Name, Len(prefixe)) = prefixe Then SousDossier = f.Path: Exit Function
    Next
    If defaut <> "" Then SousDossier = chemin & "\" & defaut
End Function

Private Function SousDossierNomme(ByVal base As String, ByVal nom As String) As String
    If fso.FolderExists(base & "\" & nom) Then
        SousDossierNomme = base & "\" & nom
    Else
        SousDossierNomme = base
    End If
End Function

Private Function SousDossierAnnee(ByVal base As String, ByVal dateMail As Date) As String
    If fso.FolderExists(base & "\" & Year(dateMail)) Then
        SousDossierAnnee = base & "\" & Year(dateMail)
    Else
        SousDossierAnnee = base
    End If
End Function

Private Sub CreerDossier(ByVal chemin As String)
    If Not fso.FolderExists(chemin) Then
        CreerDossier fso.GetParentFolderName(chemin)
        fso.CreateFolder chemin
    End If
End Sub

'==============================================================================
'  Doublons : index des noms de fichiers par dossier (sous-dossiers compris)
'==============================================================================
Private Function IndexDossier(ByVal dossier As String) As Object
    Dim dict As Object
    If cacheIndex.Exists(dossier) Then Set IndexDossier = cacheIndex(dossier): Exit Function
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    If fso.FolderExists(dossier) Then Indexer fso.GetFolder(dossier), dict
    cacheIndex.Add dossier, dict
    Set IndexDossier = dict
End Function

Private Sub Indexer(ByVal f As Object, ByVal dict As Object)
    Dim fi As Object, sf As Object
    On Error Resume Next
    For Each fi In f.Files
        dict(LCase(fi.Name)) = True
        dict("#" & CleFichier(fi.Name) & "|" & fi.Size) = True
        dict("#" & CleFichier(fi.Name)) = True
    Next
    For Each sf In f.SubFolders
        Indexer sf, dict
    Next
End Sub

Private Sub AjouterIndex(ByVal dossier As String, ByVal nom As String, ByVal taille As Long)
    Dim dict As Object
    Set dict = IndexDossier(dossier)
    dict(LCase(nom)) = True
    dict("#" & CleFichier(nom)) = True
    If taille > 0 Then dict("#" & CleFichier(nom) & "|" & taille) = True
End Sub

' Doublon = même nom, ou nom proche (date en tête, « (1) », ponctuation) ET même taille exacte
Private Function FichierDejaPresent(ByVal dossier As String, ByVal nom As String, ByVal att As Object) As Boolean
    Dim dict As Object, cle As String, tmp As String, taille As Long
    Set dict = IndexDossier(dossier)
    If dict.Exists(LCase(nom)) Then FichierDejaPresent = True: Exit Function
    cle = CleFichier(nom)
    If Not dict.Exists("#" & cle) Then Exit Function
    ' nom proche trouvé : on compare la taille exacte du fichier
    taille = TailleExacte(att, nom)
    If taille > 0 Then FichierDejaPresent = dict.Exists("#" & cle & "|" & taille)
End Function

' Taille exacte d'une pièce jointe (att.Size inclut l'enveloppe MAPI) : copie temporaire dans %TEMP% du poste,
' supprimée aussitôt. Rien n'est écrit sur le serveur.
Private Function TailleExacte(ByVal att As Object, ByVal nom As String) As Long
    Dim tmp As String
    On Error Resume Next
    tmp = Environ("TEMP") & "\archsps_" & Format(Now, "hhnnss") & "_" & nom
    att.SaveAsFile tmp
    TailleExacte = FileLen(tmp)
    Kill tmp
    On Error GoTo 0
End Function

' Clé de comparaison : sans date en tête, sans « (1) », lettres et chiffres uniquement
Private Function CleFichier(ByVal nom As String) As String
    Dim s As String, ext As String, rx As Object
    ext = LCase(fso.GetExtensionName(nom))
    s = LCase(fso.GetBaseName(nom))
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.Pattern = "^(20[0-9]{2}[ ._-]?[0-9]{2}[ ._-]?[0-9]{2})[ _-]*"
    s = rx.Replace(s, "")
    rx.Pattern = "\([0-9]+\)\s*$"
    s = rx.Replace(s, "")
    rx.Pattern = "[^a-z0-9]"
    s = rx.Replace(s, "")
    CleFichier = s & "." & ext
End Function

' Un mail est considéré comme déjà archivé si un .msg ou un .eml (script pst_archive.py) de même nom existe déjà dans 06_MAILS
Private Function MailDejaArchive(ByVal dossier06 As String, ByVal objet As String, ByVal nomMail As String) As Boolean
    Dim dict As Object, base As String
    Set dict = IndexDossier(dossier06)
    base = Left(nomMail, Len(nomMail) - 4)       ' sans « .msg »
    MailDejaArchive = dict.Exists(LCase(base & ".msg")) Or dict.Exists(LCase(base & ".eml"))
End Function

'==============================================================================
'  Outils
'==============================================================================
Private Function ExtraireAffaires(ByVal texte As String, Optional ByVal premiereSeulement As Boolean = False) As Collection
    Dim res As New Collection, mts As Object, mt As Object, num As String
    If texte = "" Then Set ExtraireAffaires = res: Exit Function
    Set mts = rxAffaire.Execute(texte)
    For Each mt In mts
        num = "7." & mt.SubMatches(1) & "." & mt.SubMatches(2)
        If CInt(mt.SubMatches(1)) >= 2014 And CInt(mt.SubMatches(1)) <= Year(Date) Then
            If Not Contient(res, num) Then res.Add num
            If premiereSeulement Then Exit For
        End If
    Next
    Set ExtraireAffaires = res
End Function

' Recherche en mots entiers : « EPHE » ne se déclenche pas sur « STEPHEN »
Private Function AffaireParCorrespondance(ByVal texte As String) As Collection
    Dim res As New Collection, paire As Variant, p As Long, t As String, k As String
    t = " " & MotsSeuls(texte) & " "
    For Each paire In Split(CORRESPONDANCES, ";")
        p = InStr(paire, "=")
        If p > 1 Then
            k = MotsSeuls(Left(paire, p - 1))
            If k <> "" Then
                If InStr(t, " " & k & " ") > 0 Then res.Add Trim(Mid(paire, p + 1)): Exit For
            End If
        End If
    Next
    Set AffaireParCorrespondance = res
End Function

' Texte normalisé réduit à ses mots (lettres et chiffres) séparés par une espace
Private Function MotsSeuls(ByVal s As String) As String
    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.Pattern = "[^A-Z0-9]+"
    MotsSeuls = Trim(rx.Replace(Normaliser(s), " "))
End Function

Private Function JoindreAffaires(ByVal c As Collection) As String
    Dim x As Variant
    For Each x In c
        If JoindreAffaires <> "" Then JoindreAffaires = JoindreAffaires & " / "
        JoindreAffaires = JoindreAffaires & CStr(x)
    Next
End Function

Private Function EstPhoto(ByVal nom As String) As Boolean
    Select Case LCase(fso.GetExtensionName(nom))
        Case "heic", "heif", "jpg", "jpeg", "png"
            EstPhoto = True
    End Select
End Function

Private Function Contient(ByVal c As Collection, ByVal v As String) As Boolean
    Dim x As Variant
    For Each x In c
        If CStr(x) = v Then Contient = True: Exit Function
    Next
End Function

Private Function AdresseExpediteur(ByVal m As Object) As String
    On Error Resume Next
    If m.SenderEmailType = "EX" Then
        AdresseExpediteur = m.Sender.GetExchangeUser.PrimarySmtpAddress
    End If
    If AdresseExpediteur = "" Then AdresseExpediteur = m.SenderEmailAddress
End Function

Private Function MailIgnore(ByVal objet As String, ByVal expediteur As String) As Boolean
    Dim e As String, o As String, v As Variant
    e = LCase(expediteur): o = LCase(objet)
    For Each v In Array("noreply", "no-reply", "mezzoteam", "resolving.com", "microsoftexchange", "wetransfer", "mailjet", "anthropic", "postmaster", "mailer-daemon")
        If InStr(e, v) > 0 Then MailIgnore = True: Exit Function
    Next
    For Each v In Array("réponse automatique", "reponse automatique", "automatic reply", "non remis", "undeliverable", "absente entre", "absent du")
        If InStr(o, v) > 0 Then MailIgnore = True: Exit Function
    Next
End Function

Private Function PJIgnoree(ByVal att As Object) As Boolean
    Dim nom As String, ext As String, cache As Boolean
    On Error Resume Next
    If att.Type = 6 Then PJIgnoree = True: Exit Function                ' olOLE
    cache = att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B")
    If cache Then PJIgnoree = True: Exit Function                       ' pièce jointe masquée (image intégrée)
    nom = LCase(att.FileName)
    ext = LCase(fso.GetExtensionName(nom))
    If nom = "" Then PJIgnoree = True: Exit Function
    If ext = "ics" Or ext = "vcf" Or ext = "p7s" Or nom Like "att0*" Then PJIgnoree = True: Exit Function
    If ext = "png" Or ext = "jpg" Or ext = "jpeg" Or ext = "gif" Or ext = "bmp" Or ext = "emz" Or ext = "wmz" Then
        If nom Like "image*" Or nom Like "outlook*" Or nom Like "logo*" Or att.Size < 60000 Then PJIgnoree = True: Exit Function
        ' image intégrée au corps du mail (signature, logo) : Content-ID ou indicateur « référencée en HTML »
        Dim cid As String, flags As Long
        cid = "": flags = 0
        cid = att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
        flags = att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37140003")
        If cid <> "" Or (flags And 4) = 4 Then PJIgnoree = True
    End If
End Function

Private Function NettoyerNom(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
        s = Replace(s, c, "_")
    Next
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    s = Trim(s)
    Do While Right(s, 1) = ".": s = Left(s, Len(s) - 1): Loop
    If s = "" Then s = "sans objet"
    NettoyerNom = s
End Function

Private Function Tronquer(ByVal dossier As String, ByVal base As String, ByVal ext As String) As String
    Dim maxi As Long
    maxi = MAX_CHEMIN - Len(dossier) - 1 - Len(ext)
    If maxi < 20 Then maxi = 20
    If Len(base) > maxi Then base = Trim(Left(base, maxi))
    ' Windows retire les points et espaces en fin de nom : on les retire aussi pour que le nom testé soit le nom écrit
    Do While Len(base) > 0 And (Right(base, 1) = "." Or Right(base, 1) = " "): base = Left(base, Len(base) - 1): Loop
    If base = "" Then base = "sans nom"
    Tronquer = base & ext
End Function

Private Function Normaliser(ByVal s As String) As String
    Dim a As Variant, b As Variant, i As Long
    s = UCase(s)
    a = Array("É", "È", "Ê", "Ë", "À", "Â", "Î", "Ï", "Ô", "Ù", "Û", "Ü", "Ç", "Œ", "œ", "Æ", "æ", "_", ".", "-", "'", "’")
    b = Array("E", "E", "E", "E", "A", "A", "I", "I", "O", "U", "U", "U", "C", "OE", "OE", "AE", "AE", " ", " ", " ", " ", " ")
    For i = 0 To UBound(a)
        s = Replace(s, a(i), b(i))
    Next
    Normaliser = s
End Function

Private Function Nz(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then Nz = "" Else Nz = CStr(v)
End Function

'==============================================================================
'  Journal
'==============================================================================
Private Sub Journaliser(ByVal m As Object, ByVal sens As String, ByVal expediteur As String, ByVal affaire As String, _
                        ByVal chemin As String, ByVal element As String, ByVal fichier As String, _
                        ByVal dest As String, ByVal statut As String)
    Dim d As String
    On Error Resume Next
    d = Format(m.ReceivedTime, "dd/mm/yyyy hh:nn")
    journal.Add Csv(d) & ";" & Csv(sens) & ";" & Csv(expediteur) & ";" & Csv(Nz(m.Subject)) & ";" & _
                Csv(affaire) & ";" & Csv(chemin) & ";" & Csv(element) & ";" & Csv(fichier) & ";" & _
                Csv(dest) & ";" & Csv(statut) & ";" & Csv(m.Parent.FolderPath)
End Sub

Private Function Csv(ByVal s As String) As String
    s = Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), """", """""")
    Csv = """" & s & """"
End Function

Private Function EcrireJournal() As String
    Dim nom As String
    nom = "Journal archivage " & Format(Now, "yyyy-mm-dd hhnnss") & " " & NettoyerNom(Environ("USERNAME")) & _
          IIf(SIMULATION, " SIMULATION", "") & ".csv"
    EcrireJournal = EcrireTexte(nom, journal)
End Function

' Écrit les lignes en UTF-8 (avec BOM) dans JOURNAL_DOSSIER, sans jamais écraser un fichier existant ;
' en cas d'échec, sur le Bureau. Renvoie le chemin écrit.
Private Function EcrireTexte(ByVal nom As String, ByVal lignes As Collection) As String
    Dim dossier As Variant, chemin As String, stm As Object, ligne As Variant
    For Each dossier In Array(JOURNAL_DOSSIER, Environ("USERPROFILE") & "\Desktop")
        On Error Resume Next
        Err.Clear
        CreerDossier CStr(dossier)
        chemin = NomLibre(CStr(dossier) & "\" & nom)
        Set stm = CreateObject("ADODB.Stream")
        stm.Type = 2: stm.Charset = "utf-8": stm.Open
        For Each ligne In lignes: stm.WriteText CStr(ligne) & vbCrLf: Next
        stm.SaveToFile chemin, 1                    ' adSaveCreateNotExist : échoue plutôt que d'écraser
        stm.Close
        If Err.Number = 0 Then EcrireTexte = chemin: Exit Function
        On Error GoTo 0
    Next
    EcrireTexte = "(journal non écrit)"
End Function

' « chemin » s'il n'existe pas, sinon « chemin (2) », « chemin (3) »...
Private Function NomLibre(ByVal chemin As String) As String
    Dim base As String, ext As String, i As Long
    NomLibre = chemin
    If Not fso.FileExists(chemin) Then Exit Function
    base = fso.GetParentFolderName(chemin) & "\" & fso.GetBaseName(chemin)
    ext = fso.GetExtensionName(chemin)
    i = 2
    Do While fso.FileExists(base & " (" & i & ")." & ext): i = i + 1: Loop
    NomLibre = base & " (" & i & ")." & ext
End Function

'==============================================================================
'  DIAGNOSTIC : liste toutes les boîtes et tous les dossiers visibles dans
'  Outlook, avec le nombre de mails de la période. Résultat dans
'  JOURNAL_DOSSIER\Diagnostic dossiers.txt
'==============================================================================
Public Sub DiagnosticDossiers()
    Dim st As Object, lignes As Collection, stm As Object, l As Variant, chemin As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    dDebut = DateSerial(CInt(Left(DATE_DEBUT, 4)), CInt(Mid(DATE_DEBUT, 6, 2)), CInt(Right(DATE_DEBUT, 2)))
    dFin = DateSerial(CInt(Left(DATE_FIN, 4)), CInt(Mid(DATE_FIN, 6, 2)), CInt(Right(DATE_FIN, 2)))
    Set lignes = New Collection
    lignes.Add "Diagnostic version " & VERSION & " - utilisateur " & Environ("USERNAME") & " - poste " & Environ("COMPUTERNAME") & _
               " - " & Format(Now, "dd/mm/yyyy hh:nn")
    lignes.Add "Boîte retenue par la macro : " & NomBoiteRetenue()
    For Each st In Application.Session.Stores
        lignes.Add ""
        lignes.Add "=== BOITE : " & st.DisplayName & "  (type " & st.ExchangeStoreType & ")"
        On Error Resume Next
        DiagDossier st.GetRootFolder, lignes, 0
        On Error GoTo 0
    Next
    chemin = EcrireTexte("Diagnostic dossiers " & NettoyerNom(Environ("USERNAME")) & " " & Format(Now, "yyyy-mm-dd hhnnss") & ".txt", lignes)
    MsgBox "Diagnostic écrit dans : " & chemin, vbInformation, "Archivage SPS"
End Sub

Private Function NomBoiteRetenue() As String
    Dim st As Object
    Set st = TrouverBoite()
    If st Is Nothing Then NomBoiteRetenue = "(aucune)" Else NomBoiteRetenue = st.DisplayName
End Function

Private Sub DiagDossier(ByVal f As Object, ByVal lignes As Collection, ByVal niveau As Integer)
    Dim sf As Object, n As Long, total As Long, filtre As String
    On Error Resume Next
    total = -1: n = -1
    total = f.Items.Count
    filtre = FiltrePeriode(dDebut, dFin)
    n = f.Items.Restrict(filtre).Count
    lignes.Add String(niveau * 2, " ") & f.Name & "  | type " & f.DefaultItemType & " | " & total & " éléments | " & n & " sur la période"
    For Each sf In f.Folders
        DiagDossier sf, lignes, niveau + 1
    Next
End Sub
