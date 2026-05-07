# paletlist

Static site for Paletlist.

## Why the site stops

If you run the site on your own PC (or on a home server connected to it), it becomes unavailable when the PC is off.

## How this repo works now

This repository is configured to deploy automatically to GitHub Pages on each push to `main` via `.github/workflows/deploy.yml`.

## One-time setup in GitHub

1. Open repository settings.
2. Go to `Pages`.
3. In `Build and deployment`, choose `Source: GitHub Actions`.
4. Push to `main` (or run the workflow manually).

After that, the site is hosted by GitHub and works independently from your PC.
