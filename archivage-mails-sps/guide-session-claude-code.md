# Démarrer une session Claude Code — Archivage des mails SPS

## 1. Préparer le poste (une seule fois)

Utilisez de préférence le poste d'une assistante, puisque c'est là que les sous-dossiers
de la boîte sont visibles.

1. **Git pour Windows** : l'installer depuis git-scm.com. Claude Code en a besoin sous Windows.
2. **Python 3.11 ou plus récent** : l'installer depuis python.org en cochant
   « Add python.exe to PATH ».
3. **Claude Code** : ouvrir PowerShell et taper :

   ```powershell
   irm https://claude.ai/install.ps1 | iex
   ```

   Fermer PowerShell, le rouvrir, puis vérifier l'installation avec `claude --version`.
4. **Copier le projet en local** : pour éviter les lenteurs du réseau, copier le dossier
   `U:\S.P.S. 26\_ARCHIVAGE MAILS\projet-archivage-sps` vers `C:\dev\projet-archivage-sps`.

## 2. Lancer la session

Dans PowerShell :

```powershell
cd C:\dev\projet-archivage-sps
git init
git config user.name "Prénom Nom"
git config user.email "prenom.nom@degouy.fr"
git add -A
git commit -m "Etat initial"
claude
```

Le premier commit enregistre l'état de départ : chaque modification faite par Claude Code
pourra ensuite être relue (`git diff`) ou annulée.

À la première connexion, choisissez le compte Claude de DEGOUY. Claude Code lit
automatiquement le fichier `CLAUDE.md`, qui contient le contexte, les règles et l'état du projet.

## 3. Premier message à copier dans Claude Code

```text
Lis CLAUDE.md et le code dans src/. Ne modifie rien pour l'instant.
1. Résume-moi en 10 lignes ce que font la macro et le script, et les règles impératives.
2. Propose-moi une architecture cible entre : (a) macro VBA améliorée sur chaque poste,
   (b) script Python lisant Outlook par COM (pywin32) sur le poste,
   (c) Microsoft Graph en lecture seule. Donne pour chacune l'effort, les risques
   et la visibilité des sous-dossiers de la boîte partagée.
3. Propose un plan en étapes courtes, avec un mode simulation en premier.
Rappel : la boîte mail ne doit jamais être modifiée et aucun fichier du serveur
ne doit être écrasé ni supprimé.
```

## 4. Pendant la session

- Validez chaque étape avant de passer à la suivante. Claude Code demande l'autorisation
  avant de modifier des fichiers ou de lancer des commandes.
- Si Claude Code demande à lancer une commande qui supprime ou déplace des fichiers
  sur `U:` ou `M:`, refusez-la.
- Pour tester, demandez un journal de simulation sur une semaine, par exemple :
  « lance une simulation du 21 au 25/09/2026 ».
- Faites relire le journal (dans `U:\S.P.S. 26\_ARCHIVAGE MAILS`) avant toute copie réelle.

## 5. Fin de session

1. Demandez à Claude Code : « mets à jour CLAUDE.md avec l'état d'avancement et les
   décisions prises ».
2. Enregistrez l'état : `git add -A` puis `git commit -m "Point du JJ/MM/AAAA"`.
3. Recopiez le dossier du projet sur `U:` **dans un nouveau dossier daté**, sans écraser
   la version précédente, par exemple :

   ```powershell
   robocopy C:\dev\projet-archivage-sps "U:\S.P.S. 26\_ARCHIVAGE MAILS\projet-archivage-sps 2026-09-25" /E
   ```

   Ne pas utiliser l'option `/MIR`, qui supprime les fichiers absents de la source.
