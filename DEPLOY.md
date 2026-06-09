# Deploy-Anleitung: LV → Plancraft auf Render.com

Du hast 4 Wege. Empfehlung: **Weg 1 (60 Sek)**.

## Weg 1: GitHub-Repo manuell (EMPFOHLEN, 60 Sek)

**Schritt 1 — GitHub-Repo erstellen:**
1. Gehe zu https://github.com/new
2. Repository name: `labi-lv2plancraft`
3. Public / Private: egal (Public ist einfacher)
4. **WICHTIG: KEINE** Checkboxen für README/.gitignore/license anklicken
5. Klicke "Create repository"

**Schritt 2 — Code hochladen:**
1. Auf der nächsten Seite siehst du "...or push an existing repository from the command line"
2. Dort steht ein Befehl wie:
   ```
   git remote add origin https://github.com/DEIN-USERNAME/labi-lv2plancraft.git
   git branch -M main
   git push -u origin main
   ```
3. Schick mir diese Zeilen + dein GitHub-Personal-Access-Token, dann führe ich das hier aus
4. **So erstellst du einen Token:** https://github.com/settings/tokens/new → "Generate new token (classic)" → Haken bei `repo` → Generate → Token kopieren (sieht aus wie `ghp_xxx...`)

**Schritt 3 — Auf Render deployen:**
1. Gehe zu https://dashboard.render.com/ (Account erstellen, falls nötig — geht mit Google-Login)
2. Klicke "New +" → "Blueprint"
3. Verbinde dein GitHub-Repo `labi-lv2plancraft`
4. Render erkennt `render.yaml` automatisch
5. Klicke "Apply" → wartet 2-3 Min → deine App ist live auf `https://labi-lv2plancraft.onrender.com`

## Weg 2: Render CLI (Schneller, 30 Sek, aber Kommandozeile)

```bash
# ZIP-Datei ist unter /tmp/labi-tool.tar.gz auf deinem Rechner
# Entpacken, in den Ordner gehen, Render CLI installieren:
cd ~
tar -xzf /tmp/labi-tool.tar.gz -C labi-tool-render
cd labi-tool-render
npm install -g @render-cli/cli  # falls noch nicht da
render login                     # Browser-Login
render blueprint launch          # erkennt render.yaml automatisch
```

## Weg 3: Render Drag&Drop (Nur falls noch verfügbar)

Render hat Drag&Drop-Deploy **in 2024 entfernt** für Webservices. Für Static Sites geht's noch, aber Flask ist kein Static Site. **Weg 1 oder 2 benutzen.**

## Weg 4: ngrok-Localhost (sofort, aber nur temporär)

```bash
# App lokal starten
cd /home/ki/hermes-workspace/labi-tool
./start.sh

# In einem zweiten Terminal: ngrok installieren + tunneln
snap install ngrok    # oder: pip install pyngrok
ngrok http 5000
# → kriegt eine URL wie https://abc123.ngrok.io
# → die funktioniert solange dein Rechner läuft
```

---

## Was du danach hast

Egal welcher Weg — am Ende läuft die App auf einer URL wie:
- `https://labi-lv2plancraft.onrender.com` (Weg 1+2)
- `https://abc123.ngrok.io` (Weg 4)

Die kannst du an Labi schicken, der lädt ein LV-PDF hoch, kriegt die Plancraft-XLSX zurück.

**Render Free-Tier-Hinweis:** Die App schläft nach 15 Min Inaktivität ein. Erster Aufruf nach dem Schlafen dauert ~30 Sek ("Service is starting up..."). Für Dauerbetrieb: 7$/Monat.
