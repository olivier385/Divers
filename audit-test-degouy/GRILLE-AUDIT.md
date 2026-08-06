# Grille d'audit — test.degouy.fr

À remplir une fois les données collectées (voir README.md). Chaque point recevra
un constat, une note (✅ conforme / ⚠️ à améliorer / ❌ problème), et une suggestion.

## 1. Technique

- [ ] Redirection http → https et cohérence apex / www / test
- [ ] Certificat TLS (émetteur, validité, chaîne)
- [ ] En-têtes de sécurité (HSTS, X-Content-Type-Options, X-Frame-Options, CSP, Referrer-Policy)
- [ ] Temps de réponse (DNS, TTFB, chargement total) et poids de page
- [ ] Compression (gzip/brotli) et mise en cache (Cache-Control, ETag)
- [ ] Version du CMS et des extensions visibles (fuites d'information : generator, readme…)
- [ ] Erreurs 404 / liens cassés internes

## 2. SEO

- [ ] Balise `<title>` et meta description uniques et pertinentes par page
- [ ] Structure des titres (un seul H1, hiérarchie logique)
- [ ] robots.txt et sitemap.xml présents et cohérents
- [ ] **Indexation de l'environnement de test** : `test.` doit être NON indexable (noindex / disallow / auth)
- [ ] Balises canoniques (risque de duplication test ↔ production)
- [ ] Open Graph / réseaux sociaux
- [ ] Données structurées (LocalBusiness pour le siège de Lognes, etc.)

## 3. Accessibilité

- [ ] Attribut `lang` sur `<html>`
- [ ] Textes alternatifs des images
- [ ] Contrastes de couleurs
- [ ] Navigation clavier et ordre de tabulation
- [ ] Libellés des formulaires (contact)

## 4. Contenu & UX

- [ ] Clarté de la proposition de valeur en page d'accueil (bureau d'études VRD / coordination SPS)
- [ ] Coordonnées visibles et exactes (adresse Lognes, téléphone, e-mail)
- [ ] Mentions légales et politique de confidentialité (RGPD, bandeau cookies)
- [ ] Adaptation mobile (viewport, mise en page responsive)
- [ ] Pages clés : présentation, savoir-faire, références/clients, recrutement, contact
- [ ] Actualité du contenu (blog, dates des dernières publications)

## 5. Spécifique environnement de test

- [ ] Différences test ↔ production (contenu, version, thème)
- [ ] Protection d'accès du site de test (le 403 observé suggère une restriction IP : bonne pratique, à confirmer)
- [ ] Données de production présentes sur le test (formulaires, e-mails réels) à vérifier
