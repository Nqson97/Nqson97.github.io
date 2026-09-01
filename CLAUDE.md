# Portfolio site — project memory

Single-file HTML/CSS/JS portfolio for **Son Nguyen Quoc** (Senior Game Designer, HCMC, Vietnam).
Live at **https://nqson97.github.io/**. Repo: `Nqson97/Nqson97.github.io` (this folder = repo root).
Second, separate repo created from this project: `Nqson97/ActionLikeSouls` (source for the "Simple UE Combat" task, see below).

## Concept

Modeled on https://www.twotinydice.com/forward/#/ — a card-game interaction:
- Projects are dealt as a fanned hand of cards at the bottom of the screen.
- Drag (or click, both supported) a card onto the play field to flip it and reveal project details.
- Each played project has a **Play** button: clicking moves that card to a small banner at the top of the field, and the hand re-deals as "task" cards representing specific work done on that project. Playing a task card shows its own detail (description + optional image/video/doc-link/YouTube embed). "Return to hand" / "◂ All Projects" navigate back out.

Everything lives in **one file**: `index.html` (~1050+ lines). No build step, no external JS/CSS libraries. Pure HTML/CSS/JS with the Pointer Events API for drag.

## Where things are (index.html)

- **`PROJECTS` array** (~line 640 onward): the entire data model. Each project: `{ id, title, type, year, art (emoji fallback), artBg (CSS gradient), img/img2 (art URLs), description, tags, link/linkLabel/link2/link2Label, tasks[] }`.
- **Task objects**: `{ title, art, img?, artBg?, description, link?, linkLabel?, video?, videoPoster?, ytEmbed? }`.
  - `link` + `linkLabel` → renders a button that opens a URL/asset in a new tab (used for PDFs and the ActionLikeSouls GitHub repo).
  - `video` + `videoPoster` → renders a self-hosted `<video>` tag.
  - `ytEmbed` → renders a `<iframe class="yt-embed">` (added for the Epic Mega Jam task). Use the `https://www.youtube.com/embed/<id>` URL form, not the watch URL.
- **`cardFrontHTML(p)`**: renders card art. Always add `referrerpolicy="no-referrer"` to any `<img>` — see Hotlink protection below.
- **`enterTaskMode`**: maps `project.tasks` into hand cards with id `project.id + '::task-' + i`. Bug fixed once: must use `t.artBg || project.artBg`, NOT `project.artBg` outright (would clobber task-level overrides).
- **Layout CSS**: hand fan math is viewport-capped (`layoutHand`, look for `maxSpan`/`gap`), header is boxed (`rgba(8,10,22,0.72)` panel), background is a purple grunge JPG (`assets/bg.jpg`, `background-attachment: fixed` on desktop). There's a `@media (min-width: 761px) and (max-height: 720px)` block that compresses the hand for short desktop windows, and a `@media (max-width: 760px)` block for phones. The `.detail` panel scrolls internally (`overflow-y: auto`) and is height-capped so long task content (video + text + buttons) can't overlap the project banner or the hand — this was a real bug the user hit and reported via screenshot once.

## Current project list

1. **Demo Work** (`demo-work`) — personal demos, always first in the hand. Tasks (in order):
   1. Creature Design — The Spikes Master (design doc PDF)
   2. Cinematics — Training Sequence (self-hosted video)
   3. FriendSlopGDD (design doc PDF, "Sea of Sheep")
   4. Epic Mega Jam 2025 (YouTube embed)
   5. Survival Game Taming System (design doc PDF, ARK-style taming)
   6. Survival Craft Recipe System (design doc PDF, tiered crafting unlocks)
   7. Simple UE Combat (links out to the `ActionLikeSouls` GitHub repo — UE 5.6, combo combat, State Tree AI, day-night cycle)
   8. Giant Drop — Full GDD (design doc PDF, vampire-survivors roguelite)
2. **Casual Puzzle & Sorting Titles** (`puzzle-titles`) — Prime Tech work, 5 tasks, no images (emoji only).
3. **Undisclosed AA Title** (`bdg-aa-title`) — 5 tasks, no images.
4. **Shooter, Survival & Boss Design** (`sparx-aa-projects`) — 5 tasks, no images.
5. **BloodLoop** (`bloodloop`) — 3 tasks, images hotlinked from bloodloop.com (see below).
6. **Mahokenshi** (`mahokenshi`) — 4 tasks, images from mahokenshi.com.
7. **Midnight Suns** (`midnight-suns`) — 4 tasks, images from assets.2k.com.
8. **SpongeBob: The Cosmic Shake** (`cosmic-shake`) — 3 tasks, images from cosmicshake.thqnordic.com.
9. **Little Hope & House of Ashes** (`dark-pictures`) — 4 tasks, images from thedarkpictures.com.

Data pulled originally from the user's CV: `SeniorGameDesign_NguyenQuocSon_CV_Eng.pdf`. Card art for the 5 CV-sourced game projects was found via web search of official title/key art and character art (all URLs verified 200 before use).

## Recurring workflow for "add a new task with this doc/video"

This has been done ~6 times now; the pattern is stable:

1. **Doc → PDF**: if given a `.doc`/`.docx`, convert via Word COM automation (no pandoc/LibreOffice installed):
   ```powershell
   $word = New-Object -ComObject Word.Application
   $word.Visible = $false
   $doc = $word.Documents.Open("<path>", $false, $true)
   $doc.SaveAs2("<outPath>.pdf", 17)   # 17 = wdFormatPDF
   $doc.Close(); $word.Quit()
   [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
   ```
   Save into `assets/docs/<slug>.pdf`.
2. **Read it for context**: `pypdf.PdfReader(...).pages[i].extract_text()` to write an accurate description (avoid guessing content). Watch for `UnicodeEncodeError` on Windows console — encode with `.encode('ascii','replace')` when printing, not needed when just using the text in code.
3. **Thumbnail**: extract an embedded image via `page.images` (pypdf) → `PIL.Image.open(io.BytesIO(im.data))` → save as `assets/<slug>-thumb.jpg`. Check dimensions across candidates and pick the most representative one (a diagram/screenshot/concept art), then **view it with the Read tool** before committing to it — don't assume from filename/size alone.
4. **Video**: for a video the user wants embedded, prefer a **YouTube embed** (`ytEmbed: 'https://www.youtube.com/embed/<id>'`) over self-hosting — costs zero repo space. Only self-host (`video`/`videoPoster` fields, into `assets/videos/`) when the user gives a raw local file with no existing YouTube upload; then re-encode with ffmpeg CRF (not framerate reduction) to fit near GitHub's 100MB soft limit, e.g. `-crf 27 -preset slow -movflags +faststart`.
5. **Edit `index.html`**: add the task object to the `Demo Work` project's `tasks[]` array (it's always been Demo Work so far).
6. **Test before deploying** — see Testing below.
7. **Commit + push** both `index.html` and any new `assets/` files.
8. **Poll the GitHub Actions deploy and verify live** — see Deployment below. Always confirm the new asset returns 200 and the task's title string appears in the live HTML before telling the user it's done.

## Testing (headless, no browser available)

The sandbox's Browser pane cannot open `nqson97.github.io` (policy-blocked) or localhost dev servers, and `file://` pages render as static snapshots (no interactivity). The working method is a **jsdom smoke test** run from a scratch directory:

- The scratchpad (`C:\Users\ASUS\AppData\Local\Temp\claude\...\scratchpad`) **does not persist across sessions** — expect to `npm init -y && npm install jsdom` again each time, and rewrite `smoke-test.js` from scratch (Edit against a stale/wiped copy will fail with "string not found" — just `Write` a fresh one).
- jsdom lacks the Pointer Capture API the site's drag handlers call — polyfill before dispatching events:
  ```js
  window.HTMLElement.prototype.setPointerCapture = function () {};
  window.HTMLElement.prototype.releasePointerCapture = function () {};
  window.HTMLElement.prototype.hasPointerCapture = function () { return false; };
  ```
- Load with `new JSDOM(html, { runScripts: 'dangerously', pretendToBeVisual: true })`, then simulate a card play by dispatching `pointerdown` + `pointerup` (bubbles: true) on the card element — this is enough to trigger the site's click-to-play path (a full drag simulation isn't necessary).
- Standard assertions per new task: hand card count after entering task mode, the new task card exists and uses the right thumbnail `src`, playing it shows the right detail text, the link/video/iframe renders with the correct `href`/`src`, and no *other* media tag leaks onto that task's detail. Use `textContent`, not innerHTML regex — `&` becomes `&amp;` in innerHTML and breaks naive regex checks (learned this the hard way once).

## Deployment (GitHub Pages via Actions workflow)

Legacy Pages "build" deployment (the default when you first enable Pages) **broke permanently** at one point — builds hung then failed with no useful error, unrelated to any real bug in the site. Fix was migrating to **workflow (Actions) deployment**:
```bash
gh api -X PUT repos/Nqson97/Nqson97.github.io/pages -f build_type=workflow
```
`.github/workflows/deploy-pages.yml` in this repo does the deploy: checkout → configure-pages → upload-pages-artifact (path `.`) → deploy-pages, with `timeout: 1800000` on the deploy-pages step (the default 10-minute timeout was hit once on a slow run even though the deploy had actually landed — 30 min avoids false failures).

Standard post-push verification loop (background bash, ~20s poll):
```bash
run=$(gh run list --workflow "deploy-pages.yml" --limit 1 --json status,conclusion)
# poll until status == "completed", then check conclusion == "success"
curl -s -o /dev/null -w "%{http_code}" https://nqson97.github.io/<new-asset-path>   # expect 200
curl -s https://nqson97.github.io/ | grep -q "<new task title>"                     # expect match
```
Each deploy re-uploads the whole repo (~tens of MB, growing with each asset) and takes a few minutes.

## Known gotchas already solved — don't re-diagnose these

- **Hotlink protection**: bloodloop.com (and potentially other project sites) return 403 when the HTTP `Referer` header is a foreign domain (verified via HEAD requests: no-referer → 200, github.io referer → 403). Fix applied globally: every generated `<img>` tag carries `referrerpolicy="no-referrer"`. This can't be done for CSS `background-image`, which is why all card art uses `<img>` tags, not CSS backgrounds, when hotlinking a third-party image.
- **Task `artBg` inheritance bug**: `enterTaskMode` must use `t.artBg || project.artBg` — using `project.artBg` unconditionally silently clobbers any task-level `artBg` override (e.g. FriendSlopGDD's custom ocean gradient, Simple UE Combat's red/black gradient).
- **Cards cut off at viewport bottom / overlapping on short windows**: fixed via the viewport-capped fan math + the `max-height 720px` media query + making `.detail` scroll internally + `.played-card` height-driven sizing (`height: min(100%, clamp(...))`) instead of fixed. If a future screenshot shows overlap again, check window height first — most of these bugs only show on short/laptop viewports, not typical desktop monitors.
- **GitHub Pages custom vs user-site URL**: this repo is the **user site** (`Nqson97.github.io`), which serves at the bare `https://nqson97.github.io/` root — no project-page subpath. Don't reintroduce a `/profolio/`-style base path anywhere.

## ActionLikeSouls repo (separate, related)

Created from `D:\Project\Test\ActionLikeSouls\` (a local Unreal Engine 5.6 project, not itself under this profolio repo). When uploading a UE project to GitHub in the future:
- **Never commit** `Binaries/`, `Intermediate/`, `Saved/`, `DerivedDataCache/` — build artifacts, huge and regenerable.
- **Exclude third-party Marketplace/Fab content folders** from the repo (check `Content/` for named packs — Paragon characters, environment/animation/sound/VFX packs, etc.) — both a size problem (was ~9.2 GB of ~9.3 GB total `Content/`) and a redistribution-license problem. Keep only the user's own Blueprints/maps/`Config`/`.uproject`, plus Epic's own default template content (e.g. the Mannequin skeletal meshes that ship with every new UE project — those are fine to include). Document required packs in a README instead of committing them.
- Repo: `Nqson97/ActionLikeSouls`, public, ~475MB, 1086 files after scoping down from the full project folder.

## Open / offered-but-undecided items

- Header "Resume" link is still a `#` placeholder — never given an actual resume file/link.
- Offered once (not accepted or declined): swapping the 🃏 field-hint glyph for ♠ (user's system renders 🃏 as a grey box/tofu).
- No attribution line added for the purple grunge background image (`https://img.magnific.com/...`) — user never asked for one, but it's sourced from a stock/free image site.
