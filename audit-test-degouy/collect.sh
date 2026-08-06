#!/usr/bin/env bash
# Collecte en lecture seule des données publiques de test.degouy.fr
# pour audit (voir README.md). Nécessite uniquement curl.
set -u

SITE="${1:-https://test.degouy.fr}"
OUT="$(cd "$(dirname "$0")" && pwd)/donnees"
mkdir -p "$OUT/pages"

log() { printf '%s\n' "$*"; }

fetch() { # fetch <url> <fichier-sortie>
  curl -sS -L --max-time 30 -A "Mozilla/5.0 (audit interne Degouy)" \
    -o "$2" -w "%{http_code} %{size_download}o %{time_total}s %{url_effective}\n" "$1"
}

log "== Collecte depuis $SITE vers $OUT =="

# 1. Page d'accueil : HTML + en-têtes + chronométrage détaillé
fetch "$SITE/" "$OUT/pages/accueil.html" | tee "$OUT/accueil.status.txt"
curl -sSI -L --max-time 30 -A "Mozilla/5.0 (audit interne Degouy)" "$SITE/" \
  > "$OUT/accueil.headers.txt" 2>&1
curl -sS -o /dev/null --max-time 30 -A "Mozilla/5.0 (audit interne Degouy)" \
  -w "DNS: %{time_namelookup}s\nTCP: %{time_connect}s\nTLS: %{time_appconnect}s\nTTFB: %{time_starttransfer}s\nTotal: %{time_total}s\nTaille: %{size_download} octets\n" \
  "$SITE/" > "$OUT/accueil.timing.txt" 2>&1

# 2. Redirections http -> https et apex/www
for u in "http://test.degouy.fr/" "https://degouy.fr/" "https://www.degouy.fr/"; do
  curl -sS -o /dev/null --max-time 20 -w "%{http_code} $u -> %{redirect_url}\n" "$u"
done > "$OUT/redirections.txt" 2>&1

# 3. Fichiers standards
for f in robots.txt sitemap.xml sitemap_index.xml favicon.ico; do
  curl -sS -L --max-time 20 -o "$OUT/$f" -w "%{http_code} /$f (%{size_download}o)\n" "$SITE/$f"
done > "$OUT/fichiers-standards.status.txt" 2>&1

# 4. Certificat TLS
echo | openssl s_client -connect "$(echo "$SITE" | sed 's|https://||'):443" \
  -servername "$(echo "$SITE" | sed 's|https://||')" 2>/dev/null \
  | openssl x509 -noout -subject -issuer -dates > "$OUT/tls.txt" 2>&1 || true

# 5. Pages internes découvertes dans l'accueil (même domaine, 15 max)
grep -oE 'href="[^"]+"' "$OUT/pages/accueil.html" 2>/dev/null \
  | sed 's/^href="//; s/"$//' \
  | grep -E "^($SITE|/)" | grep -vE '\.(css|js|png|jpe?g|svg|webp|ico|woff2?)([?#]|$)' \
  | sed "s|^/|$SITE/|" | sort -u | head -15 > "$OUT/liens-internes.txt"

n=0
while IFS= read -r url; do
  n=$((n+1))
  slug=$(echo "$url" | sed "s|$SITE/||; s|[^a-zA-Z0-9._-]|_|g; s|^$|racine|" | cut -c1-60)
  fetch "$url" "$OUT/pages/$n-$slug.html"
done < "$OUT/liens-internes.txt" > "$OUT/pages-internes.status.txt" 2>&1

log "== Terminé. Vérifiez le dossier donnees/ puis commitez-le (voir README.md). =="
