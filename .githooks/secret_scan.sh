#!/bin/sh
# Procura segredos nas linhas ADICIONADAS. Uso: secret_scan.sh <arquivo-de-diff>
# Para liberar uma linha de propósito (ex.: chave pública), inclua "secret-scan: allow" nela.
PAT='AIza[0-9A-Za-z_-]{35}|AQ\.[0-9A-Za-z_-]{20,}|-----BEGIN [A-Z ]*PRIVATE KEY-----|"private_key"[[:space:]]*:|ghp_[0-9A-Za-z]{30,}|github_pat_[0-9A-Za-z_]{30,}|APP_USR-[0-9]{6,}-[0-9A-Za-z-]{10,}|xox[abprs]-[0-9A-Za-z-]{10,}|sk-[A-Za-z0-9_-]{32,}'
HITS=$(grep -E '^\+' "$1" | grep -vE '^\+\+\+' | grep -v 'secret-scan: allow' | grep -oE "$PAT" | sed -E 's/^(.{6}).*/\1.../' | sort -u)
if [ -n "$HITS" ]; then
  echo "BLOQUEADO: possível segredo nas alterações (mostrando só o início):" >&2
  echo "$HITS" | sed 's/^/  /' >&2
  echo "Tire o segredo do código (use variável de ambiente). Se for intencional e público, marque a linha com 'secret-scan: allow'." >&2
  exit 1
fi
exit 0
