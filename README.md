# Outsourcing Ledger Manual

Static GitHub Pages publication for the outsourcing ledger input manual.

## Published files

- `index.html`
- `fig*.svg` and `fig*.webp`
- Only the PNG files under `manual-assets/` that are referenced by `index.html`

Do not publish local agent settings, Google shortcut files, browser profiles, or unreferenced screenshots.

## Update workflow

1. Edit the source manual in the working folder.
2. Copy the updated HTML to `index.html`.
3. Copy only newly referenced image files into `manual-assets/`.
4. Run the asset check workflow locally or let GitHub Actions run it.
5. Commit and push to `main`; GitHub Pages republishes the site.
