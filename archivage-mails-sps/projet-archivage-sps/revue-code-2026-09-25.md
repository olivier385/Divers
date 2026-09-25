# Revue du code — 25/09/2026

Relecture de `src/ArchivageSPS.bas` et `src/pst_archive.py` au regard des règles de `CLAUDE.md`.
**Mise à jour du 25/09/2026 : corrections apportées dans la macro v1.9 et dans `pst_archive.py`.**

| Point | État |
|---|---|
| 0 — version, nom de l'utilisateur, encodage | corrigé (v1.9 ; utilisateur et poste dans le journal et le diagnostic ; .bas en cp1252/CRLF) |
| 1.1 — écrasement des mails et pièces jointes | corrigé : existence du fichier vérifiée juste avant l'écriture ; nom sans point ni espace final |
| 1.2 / 1.3 — journal et diagnostic | corrigé : nom avec utilisateur et heure à la seconde, création seule (`adSaveCreateNotExist`), suffixe « (2) » si le nom est pris |
| 1.4 — journal du script | corrigé : ouverture en `"x"` |
| 2 — simulation qui surestime | corrigé : l'index est mis à jour aussi en simulation, avec la taille exacte |
| 3 — images, .eml, sens Envoyés, dossiers exclus, ’ | aligné entre macro et script |
| 3 — mail joint (.msg dans le mail) | **non aligné** : la macro l'enregistre, le script le signale |
| 4 — mots-clés, correspondances en mots entiers, Œ, « RESTOS » | corrigé |
| 5 — portabilité Windows du script | partiel : `racine_categorie` corrigé ; la racine du serveur reste `~/mnt` |
| Photos .heic d'un mail multi-affaires | nouveau paramètre `COPIER_PHOTOS_MULTI_AFFAIRES = False` : photo signalée une fois dans le journal, non copiée |

Le détail initial de la revue suit.

## 0. Version et format des fichiers reçus

- **La macro reçue est la v1.7**, pas la v1.8 : l'en-tête et le message de fin affichent « version 1.7 ».
  Elle **n'ajoute pas** le nom de l'utilisateur Windows au journal ni au diagnostic.
  Il faut récupérer la v1.8 avant de la lancer sur les postes des assistantes. Sinon, voir le point 2.
- **Le fichier .bas reçu était en UTF-8 avec des fins de ligne LF.** Importé tel quel dans l'éditeur VBA, les accents deviennent
  illisibles (« rÃ©ponse automatique »). Le filtre des réponses automatiques, les statuts « DÉJÀ PRÉSENT »
  et les exclusions de dossiers (« problèmes de synchronisation ») ne fonctionnent alors plus.
  Le dépôt restitue désormais le fichier en cp1252 avec des fins de ligne CRLF (voir `.gitattributes`).

## 1. Règle « ne jamais écraser » — points à corriger

| # | Fichier | Constat | Risque |
|---|---|---|---|
| 1.1 | .bas | `m.SaveAs` et `att.SaveAsFile` écrasent sans prévenir un fichier existant. La protection repose uniquement sur l'index construit en début de traitement. | Écrasement si un fichier est créé entre-temps (autre poste, autre exécution), ou si le nom écrit diffère du nom testé. Exemple : une pièce jointe sans extension devient « nom. » dans l'index, alors que Windows l'écrit sous « nom ». |
| 1.2 | .bas | Le journal est écrit avec `SaveToFile …, 2` (écrasement). Son nom est à la minute près. | Deux lancements dans la même minute : le premier journal est écrasé. |
| 1.3 | .bas | `Diagnostic dossiers.txt` porte un nom fixe et est écrasé à chaque lancement. | **Chaque poste efface le diagnostic du poste précédent.** C'est bloquant pour le plan décidé (lancement sur les postes des assistantes). |
| 1.4 | .py | `rapport` ajoute l'heure au nom du journal s'il existe déjà, puis ouvre le fichier en `"w"`. | Écrasement si deux rapports sont produits dans la même minute. Il faut ouvrir en `"x"`. |

Correction proposée (v1.9) : tester `fso.FileExists` juste avant chaque écriture et, s'il existe, journaliser « DÉJÀ PRÉSENT » ;
écrire les journaux et le diagnostic avec `SaveToFile …, 1` (création seule) ; mettre dans leur nom l'utilisateur Windows et l'heure à la seconde.

## 2. La simulation de la macro surestime les copies

Dans la macro, `AjouterIndex` n'est appelé **qu'en mode réel**. En simulation, un fichier « à copier » n'est donc pas
mémorisé. S'il revient dans un autre mail de la semaine (réponse, transfert), il est compté **une seconde fois**
« À COPIER ». En copie réelle, il serait « DÉJÀ PRÉSENT ».

Le script Python, lui, mémorise dans les deux modes. **Les journaux de simulation de la macro et du script ne sont
donc pas comparables** tant que ce point n'est pas corrigé : il suffit d'appeler `AjouterIndex` aussi en simulation.

## 3. Écarts entre la macro et le script

| Sujet | Macro .bas (v1.7) | Script .py | CLAUDE.md |
|---|---|---|---|
| Images ignorées | `image0*`, `outlook*`, < 60 Ko | `image*`, `outlook*`, `logo*`, < 60 Ko | `image*.png`, `logo*`, `outlook*` |
| Images intégrées (Content-ID) | non détectées (seul l'attribut « masqué » est testé) | détectées (0x3712 et indicateurs) | à ignorer |
| Mail joint (.msg dans le mail) | enregistré comme une pièce jointe | non extrait, signalé « à voir à la main » | — |
| Mail déjà archivé | cherche seulement le `.msg` | cherche le `.msg` et le `.eml` | — |
| Sens « Envoyés » | dossier Éléments envoyés ou expéditeur = boîte | chemin contenant « envoy » **ou « sent »**, ou expéditeur = boîte | idem macro |
| Dossiers exclus | selon le type de dossier Outlook | selon le **début du nom** (« notes », « journal », « rss »…) | — |
| Apostrophe typographique ’ | non normalisée | normalisée | — |

Conséquences concrètes :

- une fois le .py passé, la macro réenregistrera en `.msg` les mails déjà archivés en `.eml` : **le même mail sera en double** ;
- côté script, un sous-dossier nommé « Absents » ou « Présentations » (qui contiennent « sent ») sera traité comme « Envoyés », et un sous-dossier « Notes » ou « Journal » d'une assistante sera ignoré ;
- côté macro, les grandes images de signature (> 60 Ko) seront copiées dans `05`.

## 4. Classement

- Les mots-clés sont cherchés **en mots entiers** : « IC » ne se déclenche pas dans « PIC » ou « DEVIS ». Mon alerte précédente sur ce point est levée.
- `ContientMotCle` ne contient pas ICP, OUVERTURE, CSPS, PV DE REUNION ni PLAN D INSTALLATION. Une pièce jointe
  nommée « PV de réunion.pdf » est donc aussi classée d'après l'objet du mail, ce qui peut changer sa destination.
- Les `CORRESPONDANCES` sont cherchées **sans limite de mot** : « EPHE » se déclenche aussi sur « STEPHEN », par exemple.
  La première correspondance trouvée l'emporte.
- « Œ » n'est normalisé ni dans la macro ni dans le script : un objet « Restos du Cœur » ne trouve pas
  `RESTO DU COEUR`. Le pluriel « RESTOS » ne correspond pas non plus.

## 5. Le script Python ne tourne pas tel quel sous Windows

`pst_archive.py` a été écrit pour Linux : le serveur est lu dans `~/mnt`, `win()` convertit les chemins,
et `racine_categorie` découpe sur `/`. Sous Windows, `racine_categorie` renverrait le dossier cible entier :
**la recherche de doublons se limiterait au dossier final** au lieu de toute la catégorie, et des doublons seraient copiés.
Avant tout usage sur un poste, il faut paramétrer la racine (`U:\`) et utiliser `pathlib`.
Il faut aussi vérifier qu'une version de `libpff-python-ratom` s'installe sous Windows.

## 6. Piste sur les sous-dossiers invisibles

Quand la boîte n'apparaît pas comme une boîte du profil, la macro passe par `GetSharedDefaultFolder`. Cette méthode ne donne accès
qu'au dossier par défaut (Boîte de réception), **pas à ses sous-dossiers**. Les sous-dossiers ne sont visibles que si la boîte
est montée dans le profil (accès total + ajout automatique, ou « Ouvrir ces boîtes aux lettres supplémentaires »).
Le diagnostic de chaque poste indiquera dans quel cas on se trouve (ligne « Boîte retenue par la macro »).

## Ordre de correction proposé

1. Points 0, 1.3 et 2 **avant** de lancer la simulation sur les postes des assistantes. Sinon, les journaux ne sont ni exploitables ni conservés.
2. Points 1.1, 1.2 et 1.4 avant toute copie réelle.
3. Points 3 à 5 lors du passage à `regles.yaml` (règles communes aux deux implémentations), avec des tests unitaires.
