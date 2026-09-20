# Shop Rent Manager — container image for AletCloud App Hosting
#
# Why this exists: AletCloud's build auto-detection only recognizes
# Next.js / Nuxt / Node.js (via package.json) / Android (Gradle) / Java
# (Maven/.NET). This repo is a plain Python project (Flask web app +
# python-telegram-bot bot), which none of those checks match — so the
# platform falls through to a "static site" build, finds no static output
# to ship, and fails with "The platform refuses to ship an empty image".
#
# Most app-hosting platforms (AletCloud included, going by that build
# script's own multi-framework fallback logic) will use a Dockerfile at
# the repo root instead of auto-detection when one is present. Having
# this file here is what fixes the build — no Settings/Framework change
# should be needed once it's committed, but if the build still tries the
# old auto-detect path, check the app's Settings for a "Framework" or
# "Build method" option and set it to Docker/Dockerfile explicitly.
#
# This image runs server.py — bot.py's Telegram polling loop AND
# webapp.py's Flask app in one process, so both share the same rent.db
# file inside this one container (see server.py's own docstring for
# why that matters: two separate AletCloud apps/containers each get
# their own unshared copy of rent.db otherwise).

FROM python:3.12-slim

# tesseract-ocr: required by bot.py's receipt-photo reading (pytesseract
# is just a wrapper — it shells out to the real `tesseract` binary, which
# isn't a Python package and has to come from apt).
RUN apt-get update && apt-get install -y --no-install-recommends \
        tesseract-ocr \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
# waitress is already in requirements.txt and is what _run_web() in
# webapp.py uses as the production WSGI server — no separate gunicorn
# needed.
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# rent.db and the uploads/ folder should be treated as persistent data,
# not baked into the image. If AletCloud offers a persistent volume /
# disk mount for this app, mount it at /app so rent.db and uploads/
# survive redeploys and restarts — otherwise every new deploy starts
# from whatever rent.db happened to be in the repo at build time, and
# anything written while the container was running is lost when it's
# replaced.

# webapp.py reads WEB_PORT (default 5000); AletCloud very likely injects
# the port to listen on as $PORT (the near-universal convention for this
# kind of platform). This maps one to the other at container start so
# whichever the platform actually sets, the app binds to it.
ENV WEB_HOST=0.0.0.0
EXPOSE 5000

# If a persistent volume gets mounted at /app (see the note above), it can
# start out empty and wipe out the uploads/ folder that was baked into the
# image. Recreate it at container start (harmless if it already exists) so
# receipt uploads never fail just because the folder is missing.
CMD ["sh", "-c", "mkdir -p uploads; export WEB_PORT=${PORT:-${WEB_PORT:-5000}}; python server.py"]
