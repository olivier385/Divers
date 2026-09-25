# Analyse des journaux de simulation — 25/09/2026

Période simulée : du 21 au 25/09/2026. Deux journaux comparés :

- **Macro v1.8** sur le poste d'Olivier (`o.degouy`), lancée à 17 h 29, avec son diagnostic ;
- **Script PST** sur un nouvel export .pst, lancé à 15 h 15.

## 1. Les sous-dossiers sont désormais visibles

Le diagnostic et le journal de la macro v1.8 montrent **91 dossiers**, dont tous les sous-dossiers de la Boîte de réception :
Sophie (CR/PPSPS/DIVERS, CORDO et ses coordonnateurs), Aurore\Classement, Audrey, Heures. Le nouvel export .pst les contient aussi.
**Le problème des « sous-dossiers invisibles » est donc résolu sur le poste d'Olivier**, avec la macro comme avec l'export.
Lancer la macro sur les postes des assistantes n'est plus indispensable : ce serait seulement une vérification.

Dossiers qui ont reçu des mails sur la semaine : Boîte de réception (50), Sophie\CORDO\Christelle (34), Aurore\Classement\Amélie (22),
Sophie\CR/PPSPS/DIVERS (24), Sophie\CORDO\Franck (19), Aurore\Classement\Charlotte (13), Sophie\CORDO\Jean (9), Thierry (4)…
Les dossiers d'Audrey n'ont reçu aucun mail sur la période.

## 2. Anomalie : la macro ignore la matinée du premier jour

La macro ne voit **aucun mail reçu le lundi 21/09 avant 12 h 00**. Son premier mail date de 12 h 01, alors que l'export .pst en compte 7 avant midi.

**Cause** : le filtre de date d'Outlook était écrit au format `h:nn AMPM`. Sur un Windows français, sans indicateur AM/PM,
minuit s'écrit « 12:00 », et Outlook le comprend comme midi. **Corrigé en v1.10** : le filtre est écrit en 24 h, sur une plage élargie
d'un jour de chaque côté, puis la date exacte est vérifiée mail par mail.

C'est ce qui explique pourquoi la macro ne signalait pas les 10 photos .heic du mail « Oudiné » (7.2023.406 / 407 / 2026.045),
reçu le 21/09 à 10 h 51.

## 3. La simulation v1.8 compte deux fois les mails présents en double dans la boîte

| | Macro v1.8 | Script PST |
|---|---:|---:|
| Mails « à enregistrer » | 377 | 288 |
| Mails « déjà présents » | 0 | 102 |
| Noms de mails distincts à enregistrer | **278** | 288 |
| Pièces jointes « à copier » | 91 | 88 |
| … dont pièces jointes distinctes | **63** | 88 (dont 30 photos .heic = 10 × 3 affaires) |
| Pièces jointes déjà présentes | 282 | 320 |
| Lignes « sans n° d'affaire » | 20 | 20 |

Beaucoup de mails de la semaine sont **à la fois dans un sous-dossier et dans Éléments supprimés**
(par exemple « PDM_RJ 260915-16 », dans Christelle et dans Éléments supprimés).
La macro v1.8 les compte deux fois « à enregistrer » : 99 mails et 28 pièces jointes en trop.
En copie réelle, le second exemplaire serait retrouvé comme « déjà présent ». La v1.9 corrige ce comptage.

Une fois ces deux défauts écartés, les deux outils voient **les mêmes mails** : 289 en commun. Les seuls écarts sont
les 7 mails de lundi matin et un mail horodaté 17:05 dans le .pst et 17:06 dans Outlook.

**Ordre de grandeur réel pour la semaine** : environ 285 mails à enregistrer et 60 pièces jointes à copier
(hors photos d'un mail qui cite plusieurs affaires), sur une quarantaine d'affaires.

Répartition des 63 pièces jointes distinctes (macro) : RJ 35 · PPSPS/PIC 16 · autres documents 14 · CR de maîtrise d'œuvre 13 · IC 8 · PGC 2 · CISSCT, DIUO, avis APS/AVP 1 chacun
(les chiffres portent sur les 91 lignes avant dédoublonnage).

## 4. Mails sans n° d'affaire (17 distincts)

Ils se répartissent en deux groupes.

- **Commerciaux** (normal qu'ils n'aient pas de n° d'affaire) : consultation ZAC des Deux Moulins, demandes de devis
  Paris Sud Aménagement et CD93, autorisation de commande, facturation résiduelle CD91.
- **Chantiers, à ajouter à `CORRESPONDANCES`** (n° d'affaire à fournir) :
  - Impasse des Chantereines à Montreuil (3 mails : Artelia, Est Ensemble) ;
  - SEQENS St-Denis Gai Logis (CR n° 122) ;
  - Épinay-sur-Seine (CR 01) ;
  - Sciences Po (rendu APD) ;
  - SIAAP Dalle (Razel) ;
  - DIUO La Montagne Pierreuse Montreuil ;
  - Requalification terrain de football (atelier Chaneac) ;
  - PNEI (Setec) ;
  - Convention de prêt d'échafaudage (Ascistego) ;
  - Réunion de chantier du 23/09 (Tersen).

## 5. Point de vigilance entre les deux outils

Un même mail peut être horodaté à une minute d'écart par Outlook et par l'export .pst (17:05 / 17:06).
Son nom de fichier diffère alors (`2026-09-21 1705 - …eml` contre `… 1706 - ….msg`), et il serait archivé deux fois
si l'on utilisait successivement le script puis la macro sur la même période. **Il faut choisir un seul outil par période.**

## Suite proposée

1. Importer la **v1.10** sur le poste d'Olivier et relancer la simulation du 21 au 25/09. On attend environ
   285 mails à enregistrer, 60 pièces jointes à copier et 10 photos signalées.
2. Compléter `CORRESPONDANCES` avec les n° d'affaire listés au point 4.
3. Valider le journal, puis passer `SIMULATION = False` pour la copie réelle de la semaine.

## Résultat de la simulation v1.10 (poste o.degouy, 25/09/2026 à 17 h 37)

| | Attendu | Obtenu |
|---|---:|---:|
| Mails à enregistrer | ~285 | **290** |
| Pièces jointes à copier | ~60 | **66** |
| Photos de mails multi-affaires | 10 | **10** |
| Mails sans n° d'affaire | 20 | 20 |
| Erreurs | 0 | 0 |

Autres chiffres : 67 dossiers parcourus, 403 mails traités, 380 pièces jointes examinées, 416 fichiers déjà présents, durée 1,3 min.
Journal : `Journal archivage 2026-09-25 173745 o.degouy SIMULATION.csv`.
Les résultats sont conformes à ce qui était attendu. Reste à relire le journal ligne par ligne avant la copie réelle.
