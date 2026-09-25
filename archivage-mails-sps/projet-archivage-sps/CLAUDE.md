# Archivage des mails SPS — Groupe DEGOUY

Contexte à lire par Claude Code au début de chaque session.

## Objectif

Archiver automatiquement sur le serveur les mails de la boîte partagée **secretariat.sps@degouy.fr** et leurs pièces jointes, dans le dossier de chaque affaire. Le but est que chaque livrable SPS (RJ, IC, PPSPS, PGC, DIUO, CISSCT…) soit retrouvable et **facturable**.

Les compte-rendus de réunion de maîtrise d'œuvre sont classés **systématiquement** dès que DEGOUY est destinataire.

## Règles impératives (non négociables)

1. **La boîte mail reste intacte.** Aucun mail n'est supprimé, déplacé, marqué lu ou modifié. On ne fait que **lire et copier**. N'utiliser aucune fonction Outlook d'archivage ou de déplacement, ni aucune API qui écrit dans la boîte.
2. **Ne jamais écraser un fichier du serveur.** Si le nom existe déjà, on ne touche à rien. Écrire en mode exclusif (`open(..., "xb")`).
3. **Ne jamais supprimer de fichier sur le serveur.**
4. **Simulation d'abord.** Chaque traitement a un mode simulation (par défaut) qui produit uniquement un journal CSV. La copie réelle n'a lieu que sur validation explicite d'Olivier.
5. Un .pst sur le serveur s'ouvre en **lecture seule**.
6. Ne jamais envoyer de mail. Les brouillons restent des brouillons.

## Serveur

- Lecteur `U:\` : dossiers `S.P.S. 14` à `S.P.S. 26` (un par année d'ouverture de l'affaire).
- Dossier d'une affaire : `U:\S.P.S. AA\7.AAAA.NNN <libellé>`. Exemple : `U:\S.P.S. 26\7.2026.272 …`.
- Numéro d'affaire (expression régulière) : `(^|[^0-9])7[ ._-]{0,3}(20[0-9]{2})[ ._-]{1,3}([0-9]{3})(?![0-9])`, années 2014 à l'année en cours.
- Sous-dossiers types. On cherche le préfixe numérique ; les libellés varient d'une affaire à l'autre.

| Préfixe | Dossier |
|---|---|
| 01 | `01_RJ et IC`, avec les sous-dossiers `RJ` et `IC`, eux-mêmes parfois découpés par année (`2026`) |
| 02 | `02_PPSPS et PIC` |
| 03 | `03_CR CHANTIER` |
| 04 | `04_CISSCT` |
| 05 | `05_DOCS TECH` ou `05_DIVERS` |
| 06 | `06_MAILS\Reçus` et `06_MAILS\Envoyés` |
| 07 | `07_DIUO` |
| 09 | `09_PGC` |
| 10 à 13 | `10_APS et AVP`, `11_APD`, `12_PRO`, `13_DCE` |
| 16 | `16_DOSSIER COMPTABLE` / `16_Devis Contrats` |
| 17 | `17_Revue de contrat` |

- Dossier de travail et journaux : `U:\S.P.S. 26\_ARCHIVAGE MAILS`.

## Règles de classement (reprises de la macro v1.8)

**Rattachement d'un mail à une affaire**, dans cet ordre :

1. le numéro d'affaire dans l'objet ;
2. sinon la table `CORRESPONDANCES` (mot-clé de l'objet ou de l'adresse de l'expéditeur vers le numéro d'affaire) ;
3. sinon le premier numéro trouvé dans les 3 000 premiers caractères du corps.

**Enregistrement du mail** dans `06_MAILS\Reçus` ou `\Envoyés`, sous le nom `aaaa-mm-jj hhmm - <objet nettoyé>.msg` (ou `.eml`). Le mail est « Envoyés » si l'expéditeur est la boîte elle-même ou s'il vient du dossier Éléments envoyés.

**Classement des pièces jointes.** On teste d'abord le nom du fichier, puis l'objet du mail si le nom ne contient aucun mot-clé. Les règles s'appliquent dans cet ordre :

| Mots-clés | Destination |
|---|---|
| PPSPS / PIC | 02 |
| AVIS + PRO / DCE / APD / APS-AVP | 12 / 13 / 11 / 10 |
| DIUO | 07 |
| CISSCT | 04 |
| PGC | 09 |
| LIVRET (d'accueil) | 05 |
| RJ / REGISTRE | 01\RJ\année |
| IC / ICMOD / VIC / ICP / INSPECTION | 01\IC\année |
| CR, COMPTE RENDU ou PV DE REUNION, **mail externe**, sans SPS/CSPS | 03 (CR de maîtrise d'œuvre) |
| VI / VIxxxx / VISITE / CR SPS / CSPS / OUVERTURE / CRxx | 01\RJ\année |
| DCE / APD / APS-AVP seuls | 13 / 11 / 10 |
| Sinon | 05 |

- **Mail « externe »** : expéditeur hors @degouy.fr, ou objet commençant par TR:/FW:/Fwd:.
- **Pièce jointe qui cite une autre affaire dans son nom** : elle est rangée dans cette autre affaire.
- **Pièces jointes ignorées** :
  - pièces jointes masquées ou intégrées (Content-ID), objets OLE ;
  - fichiers .ics, .vcf, .p7s, ATT0000x ;
  - images de moins de 60 Ko, image*.png, logo*, outlook*.
- **Doublons.** Un fichier est considéré comme déjà présent si :
  - le même nom existe n'importe où sous le dossier de catégorie (par exemple `01_RJ et IC`) ;
  - ou une clé normalisée identique existe avec la même taille exacte. Clé normalisée : date en tête retirée, « (1) » retiré, lettres et chiffres uniquement.
- **Expéditeurs ignorés** : noreply, mezzoteam, wetransfer, mailjet, postmaster, mailer-daemon… **Objets ignorés** : réponses automatiques, non remis.
- **Journal CSV** (UTF-8 avec BOM, séparateur `;`), colonnes : Date;Sens;Expéditeur;Objet;Affaire;Dossier affaire;Élément;Fichier;Destination;Statut;Dossier Outlook.

## Code existant (dossier `src/`)

- `ArchivageSPS.bas` : macro Outlook VBA v1.10 (v1.9 : corrections de la revue du 25/09/2026, voir `revue-code-2026-09-25.md` ; v1.10 : filtre de date en 24 h, la matinée du premier jour était ignorée, voir `analyse-journaux-2026-09-25.md`). Journal : `Journal archivage AAAA-MM-JJ hhmmss <utilisateur> [SIMULATION].csv` ; diagnostic : `Diagnostic dossiers <utilisateur> AAAA-MM-JJ hhmmss.txt`. Paramètre `COPIER_PHOTOS_MULTI_AFFAIRES` (False par défaut). Stockée en UTF-8 dans git, restituée en cp1252/CRLF, en liaison tardive (late binding), encodée en cp1252 avec des fins de ligne CRLF. Elle fonctionne sur Boîte de réception, Éléments envoyés et Éléments supprimés. Les paramètres sont en tête de fichier (dates, `SIMULATION`, `CORRESPONDANCES`).
- `pst_archive.py` : mêmes règles que la macro v1.9, appliquées à un export .pst (le serveur est lu dans `~/mnt` : script prévu pour Linux), avec pypff (paquet `libpff-python-ratom`). Il écrit les mails en .eml. Il traite par tranches avec reprise (`~/pstrun`) et propose les commandes `tree`, `run` et `report`.

## Problèmes connus

- **Revue du code du 25/09/2026** : voir `revue-code-2026-09-25.md` (écrasements possibles, simulation de la macro qui surestime les copies, écarts entre la macro et le script, script non portable sous Windows).

- **Sous-dossiers de la Boîte de réception** (Heures, Aurore, Audrey, Sophie\CORDO\…, Aurore\Classement\…) : **résolu le 25/09/2026**. La macro v1.8 et le nouvel export .pst les voient depuis le poste d'Olivier (91 dossiers).
- **Mails en double dans la boîte** : beaucoup de mails sont à la fois dans un sous-dossier et dans Éléments supprimés. Le nom de fichier (date + objet) évite de les archiver deux fois.
- **Un seul outil par période** : Outlook et l'export .pst peuvent horodater un même mail à une minute d'écart ; le nom de fichier diffère alors entre .msg et .eml.
- **Export .pst sur le lecteur réseau U:** : erreurs « inconnues » et fichier verrouillé tant qu'Outlook est ouvert. Il faut exporter en local, puis copier.
- **Réimport du .bas** : réimporter le fichier crée « ArchivageSPS1 ». Il faut d'abord supprimer les anciens modules.
- **Photos de téléphone (.heic, .jpg)** jointes à des mails qui citent plusieurs affaires : elles seraient copiées plusieurs fois. À traiter ; proposition : signaler dans le journal plutôt que copier.
- **Mails sans numéro d'affaire** : environ 15 par semaine. Il faut enrichir `CORRESPONDANCES`.
- **Archives .pst historiques** (lecture seule), dans `M:\SPS\secretariat.sps` :
  - `archives secretariatsps - Audrey.pst` : 4,7 Go, 2023 à sept. 2025 ;
  - `Aurore.pst` et `Sophie.pst`.

## Chantier proposé pour la session de codage

1. Choisir la cible : garder la macro VBA sur chaque poste, ou passer à un script Python sur Windows qui lit Outlook par COM (`pywin32`). Critères : installation sur les postes, maintenance, visibilité des sous-dossiers. Autre option : Microsoft Graph (en lecture seule), si l'informatique crée une application Azure avec une permission `Mail.Read` sur la boîte partagée.
2. Sortir les règles (mots-clés, correspondances, dossiers ignorés) dans un fichier de configuration unique (`regles.yaml`), commun à toutes les implémentations.
3. Écrire des tests unitaires du classement à partir des journaux existants (`Journal archivage … SIMULATION.csv` dans `_ARCHIVAGE MAILS`).
4. Ajouter un rapport de synthèse par affaire : livrables archivés dans la semaine, à rapprocher de la facturation.
5. Lancer une simulation sur la semaine du 21 au 25/09/2026, la faire valider par Olivier, puis faire la copie réelle. Étendre ensuite à janvier 2026.

## Conventions

- Langue : français pour les messages, les journaux et les commentaires.
- Chemins Windows ; ne pas supposer de disque local autre que `C:`.
- Toute nouvelle version incrémente le numéro affiché dans le message de fin.
