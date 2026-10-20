#!/bin/bash
# Équivalent macOS de START_GUI.bat : lance Chrome avec le port de débogage 9222.
# Usage : ./start_chrome_mac.sh   (la 1re fois : chmod +x start_chrome_mac.sh)
PORT="${SCRAPER_CHROME_PORT:-9222}"
CHROME="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
PROFILE="$HOME/Library/Application Support/Google/Chrome/Debug Profile"

[ -x "$CHROME" ] || { echo "Google Chrome introuvable dans /Applications"; exit 1; }

# Déjà lancé avec le port de débogage ? Rien à faire.
if curl -s --connect-timeout 2 "http://127.0.0.1:$PORT/json/version" >/dev/null; then
  echo "Chrome écoute déjà sur le port $PORT"; exit 0
fi

# Chrome ouvert sans port de débogage : il faut le fermer (sinon l'option est ignorée)
if pgrep -x "Google Chrome" >/dev/null; then
  echo "Fermeture de Chrome (relance avec le port $PORT)..."
  pkill -x chromedriver 2>/dev/null
  osascript -e 'quit app "Google Chrome"'
  for _ in $(seq 1 10); do pgrep -x "Google Chrome" >/dev/null || break; sleep 1; done
fi

rm -f "$PROFILE/SingletonLock" "$PROFILE/SingletonSocket" "$PROFILE/SingletonCookie"

# Profil dédié obligatoire (Chrome 136+ ignore le port sur le profil par défaut)
nohup "$CHROME" \
  --remote-debugging-port="$PORT" \
  --user-data-dir="$PROFILE" \
  --disable-blink-features=AutomationControlled \
  --no-first-run --no-default-browser-check --no-restore-last-session \
  >/dev/null 2>&1 &

for _ in $(seq 1 30); do
  curl -s --connect-timeout 2 "http://127.0.0.1:$PORT/json/version" >/dev/null && { echo "Chrome prêt sur le port $PORT"; exit 0; }
  sleep 1
done
echo "Le port $PORT ne répond pas"; exit 1
