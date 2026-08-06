# Audit de test.degouy.fr

Ce dossier contient le nécessaire pour auditer le site **https://test.degouy.fr/** (analyse et suggestions).

## Contexte

Le site `test.degouy.fr` (ainsi que `degouy.fr` et `www.degouy.fr`) renvoie **403 Forbidden**
aux requêtes provenant des environnements cloud de Claude Code : impossible de l'analyser
directement depuis une session distante. Il faut donc collecter les données depuis une
machine autorisée (la vôtre), les pousser dans ce dépôt, puis laisser Claude faire l'analyse.

## Mode d'emploi

1. Sur votre machine (Mac, Linux, ou Windows avec Git Bash / WSL), exécutez :

   ```bash
   cd audit-test-degouy
   bash collect.sh
   ```

   Le script ne fait que des requêtes HTTP en lecture (curl) vers le site et écrit
   les résultats dans `audit-test-degouy/donnees/`.

2. Vérifiez rapidement le contenu de `donnees/` (rien de confidentiel n'y est collecté :
   uniquement des pages publiques, en-têtes HTTP, robots.txt, sitemap…).

3. Commitez et poussez :

   ```bash
   git add audit-test-degouy/donnees
   git commit -m "Ajout des données collectées sur test.degouy.fr"
   git push
   ```

4. Demandez à Claude de reprendre l'analyse : il produira le rapport complet
   (technique, SEO, accessibilité, contenu, suggestions) dans ce dossier.

## Contenu du dossier

| Fichier | Rôle |
|---|---|
| `collect.sh` | Script de collecte à exécuter depuis une machine autorisée |
| `GRILLE-AUDIT.md` | Grille d'audit qui sera remplie lors de l'analyse |
| `donnees/` | (créé par le script) Pages HTML, en-têtes, robots.txt, sitemap, mesures |
