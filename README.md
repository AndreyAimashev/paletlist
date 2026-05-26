# Paletlist

Paletlist is a warehouse/order management web app. The repository contains a
static browser UI, a Python API server, SQLite-backed data storage, and PDF/XLSX
generators for packing and pallet sheets.
Production deployment is
performed to a server over SSH by GitHub Actions.

## Repository layout

- `index.html` - login page.
- `admin.html` - main single-page admin UI.
- `api_server.py` - Python HTTP API on `127.0.0.1:8081`.
- `nomenclature.json` - seed data used when the SQLite database is empty.
- `packing_sheets_generic.py` and `packing_sheet_generic_styles.css` - generic
  packing sheet HTML/PDF generation.
- `packing_sheets_lab_industries.py` - Lab Industries specific pallet sheet PDF
  generation.
- `pallet-sheet-template-generic.html` - HTML template for generated sheets.
- `requirements.txt` - Python dependencies for PDF, barcode, and XLSX support.
- `deploy/paletlist-api.service.example` - example systemd unit for the API.
- `.github/workflows/deploy.yml` - SSH deployment workflow.

## Runtime architecture

Production is expected to run as:

1. nginx serves the static files from the repository checkout.
2. nginx proxies `/api/` requests to `api_server.py` on `127.0.0.1:8081`.
3. `api_server.py` stores runtime data in `warehouse.db` next to the code.
4. systemd manages the API process via `paletlist-api.service`.

`warehouse.db` is created at runtime and should be backed up on the server if the
data is important.

## Local development

Install Python dependencies:

```bash
python3 -m pip install -r requirements.txt
```

Run the API server:

```bash
python3 api_server.py
```

The API listens on `127.0.0.1:8081`. For a browser setup that matches
production, serve the HTML files through nginx or another local static server
and proxy `/api/` to the Python process.

PDF generation with WeasyPrint also needs system libraries. The production
workflow installs:

```bash
sudo apt-get install -y \
  libpango-1.0-0 libpangocairo-1.0-0 libpangoft2-1.0-0 libcairo2 \
  libharfbuzz0b libharfbuzz-subset0 libfontconfig1 libgdk-pixbuf-2.0-0 \
  libglib2.0-0 shared-mime-info
```

Optional TTF fonts can be placed in `fonts/`; the repository ignores font files.

## Server setup

The deployed checkout is expected at:

```text
/var/www/paletlist
```

Basic one-time server setup:

1. Clone this repository into `/var/www/paletlist`.
2. Install Python and the dependencies from `requirements.txt`.
3. Configure nginx to serve the static files and proxy `/api/` to
   `127.0.0.1:8081`.
4. Install and adjust `deploy/paletlist-api.service.example` as
   `/etc/systemd/system/paletlist-api.service`.
5. Start and enable the service:

   ```bash
   sudo systemctl daemon-reload
   sudo systemctl enable --now paletlist-api.service
   ```

The nginx configuration is server-specific and is not stored in this repository.

## Deployment

Pushes to `main` run the `Deploy to server` workflow in
`.github/workflows/deploy.yml`.

The workflow connects to the server over SSH using these GitHub Actions secrets:

- `SERVER_HOST`
- `SERVER_USER`
- `SERVER_SSH_KEY`

On the server it:

1. changes to `/var/www/paletlist`;
2. checks out and pulls `main`;
3. verifies that key API routes/functions are present;
4. installs Python dependencies and required WeasyPrint system libraries;
5. reloads nginx;
6. restarts `paletlist-api.service`.

If deployment fails, check the GitHub Actions run log and the server-side
systemd/nginx logs.
