---
title: mamaXO Document Factory
emoji: 📄
colorFrom: orange
colorTo: gray
sdk: docker
app_port: 7860
pinned: false
---

# mamaXO Document Factory

Internal tool. Colleagues sign in with a shared password, fill a short form, and
download a branded PDF (Golden Visa or Property Management). The design is locked
in `base.css` + `brand.config.json`; users only supply content.

Set two secrets on the host: `APP_PASSWORD` (team password) and `SECRET_KEY` (any
random string). See `DEPLOY-GUIDE.md`.
