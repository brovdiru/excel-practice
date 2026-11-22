# Job Test Drills

Free practice tools for pre-employment skills tests. Static site, served by GitHub Pages.

- `/` — hub page (`src/hub.html`)
- `/excel/` — Excel formula trainer and topic pages (`src/app.html`, page texts in `tools/build.py`)
- Old root URLs of the Excel pages are redirect stubs to `/excel/…`.

## Build

Edit the sources, then run `python3 tools/build.py`. When a `CNAME` file exists (GitHub adds it after a custom
domain is connected), canonical URLs, `robots.txt` and `sitemap.xml` use that domain; otherwise the github.io address.

## Adding a trainer

Put it in its own folder (for example `/typing/`), add a card to `src/hub.html` and its URLs to the sitemap in `tools/build.py`.
