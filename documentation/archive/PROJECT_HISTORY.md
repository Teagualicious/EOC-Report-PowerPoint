> **ARCHIVED - NOT AUTHORITATIVE.** This file is preserved as project history. Use `../MODEL_HANDOFF.md` and `../CURRENT_ARCHITECTURE.md` for the current implementation.

# Ingestion Engine — Project State & Continuation Guide

## v16.3 — Beta 1.3 Field-Test Fixes (July 2026)

### Bug Fixes
- **Formats silently never applied — anywhere** (`engine/pptx_formats.py`,
  `engine/metrics_catalog.py`, `mapper/format_popup.py`,
  `mapper/sidebar.py`, `mapper/slide_view.py`, `engine/pptx_fill.py`):
  pandas aggregations return **numpy scalars**, and `numpy.int64` does
  NOT subclass Python `int` — so every `isinstance(v, (int, float))`
  gate failed for integer-summed metrics and values fell through to raw
  unformatted strings. This single root cause produced two reported
  symptoms: the Format Cells preview ignoring the selected options, and
  formats not carrying into the PowerPoint. Fixes: `_coerce_number()`
  converts numpy scalars at every formatting sink, `metrics_catalog`
  emits native Python numbers at the source, and a shared
  `format_with_details()` is now the single formatter used by the popup
  preview, the sidebar, live-preview insertion, and `fill_template` —
  the preview now shows exactly what will be inserted.
- **Detailed formats now travel with assignments**: decimals, commas,
  prefix and suffix were stored per-metric in the mapper session but
  never written into the mapping — saved templates and Auto-Fill only
  saw the coarse format string. Assignments now carry a
  `format_details` dict which the live preview, slide-navigation
  re-apply, and both fill paths honor.
- **Clear All destroyed multi-line formatting** (`engine/pptx_live.py`):
  the paragraph snapshot includes trailing `\r` marks; restoring them
  verbatim inserted extra paragraph breaks, shifted indices mid-loop,
  and cascaded paragraph 1's formatting (large green font) across the
  whole box. `_set_paragraph_text` now writes through a
  `Characters(1, content_len)` range that **excludes the paragraph
  mark** — structurally incapable of merging or splitting paragraphs —
  and both assignment and restore use it (restore strips marks first).
- **Non-contiguous column selection** (`mapper/query_builder.py`): the
  export-columns listbox used `selectmode="extended"` (Ctrl+click
  required for gaps); now `"multiple"` — each click toggles a row, so
  picking the first and last columns takes two plain clicks.

### Found by end-to-end pipeline testing (stand-in template + synthetic data)
- **Date/case detection scoped to the replace target** (`engine/pptx_fill.py`,
  `mapper/slide_view.py`): partial replaces detected date style and text
  case from the WHOLE text frame, so an unrelated placeholder later in
  the line (e.g. "Pulled [MONTH/YEAR]") hijacked the format — a date
  range aimed at "[XX/XX/XXXX-XX/XX/XXXX]" rendered as "June 2026".
  Detection now runs against the replace target itself. Side benefit:
  case matching now keys off the target too, so replacing an all-caps
  "CLIENT NAME" placeholder correctly uppercases the client name.
- **XX date placeholders now detectable** (`metric_dictionary.json`):
  added `X{2}/X{2}/X{4}` patterns so "[XX/XX/XXXX]"-style placeholders
  detect as us_slash and date ranges render as "06/01/2026 – 06/30/2026"
  instead of falling through to raw ISO.

### v16.18 — Excel field fixes: overlay at 0,0, dual files, dashboard look
- ActiveX txtSearch landed at 0,0 covering the banner (the "white bar
  on A1"): pywin32 unreliably passes named position args to
  OLEObjects.Add — position/size now set via properties AFTER creation,
  Placement=xlMoveAndSize.
- Dual .xlsx + .xlsm output: the post-SaveAs xlsx delete can hit a
  transient lock (AV/indexing) and failed silently. Now retried 6× with
  backoff + warning; write_to_excel also prefers an existing .xlsm and
  deletes a stale same-name .xlsx at export start, so the pair
  self-heals on the next run.
- Workbook_Open handler injected into ThisWorkbook: clears the
  "enable macros" notice the moment the workbook opens with macros on.
- Dashboard-not-spreadsheet pass (coordinates unchanged): row/column
  headers hidden, full white canvas painted A1:T80, banner/accent/rule
  span the canvas, taller banner + input rows.
Deferred still (the deck's remaining Excel look): attached suggestions
dropdown, removable term chips, auto KPI strip.

### v16.17 — HOTFIX: RoundButton double-fire (field bug)
RoundButton bound <Button-1> on the canvas widget AND on the drawn
shape/text tags — every click over the button fired the command TWICE.
Field symptoms explained: modal browse dialogs reopening after file
selection (second fire queued behind the modal), and app-wide lag
(every action executing twice — parses, window builds). Fix: single
widget-level binding. Plus a shared tkfont cache (one Font per style
tuple, not per button). Regression test now uses event_generate REAL
clicks (center + corner + disabled) — the original test called invoke()
directly and missed the double binding.

### v16.16 — Completion Rate: Σcompletions ÷ Σimpressions (approved)
Per John's rule, completion rate is now DERIVED, not averaged — naive
averaging overweighted small campaigns (testkit: naive 0.83 vs true
56.37%). Numerator preference: Completions/Contributions display-merge
(kpi display layer already unified them), then 100% Completions.
- engine/kpi.py: _derive_completion_rate overrides the averaged value
  for totals AND per-campaign details; values are percentages.
- Excel: _Config marks Completion Rate/Percent agg="rate"; modSearch
  accumulates Impressions + the completions count (PickNumerator) even
  when not displayed, riding the existing source-collapse via Chr(4)
  pseudo-metric keys; WriteResults renders 100*num/imp as "0.00%"
  with graceful fallback to the averaged value when denominators are
  missing (e.g. rate-only Architect rows without paired impressions —
  the impressions-weighted variant remains a future option there).
141 tests; VBA structure linted; math port verified against a
generated workbook.

### v16.15 — Full UI remake ("big surgery"): rounded buttons, fonts, layout
**RoundButton** (ui/utils): canvas-drawn smooth rounded-rect button —
tkinter has no native rounded corners. tk.Button-compatible API (text,
command, font, bg, fg, padx, pady, state; other tk.Button kwargs
accepted+ignored), hover shade derived from bg automatically, disabled
state, config/configure/cget/invoke passthrough, pack/grid/place
delegation. Corner illusion via parent-bg canvas. ALL 54 tk.Buttons
across ui/* and mapper/* converted mechanically (grep confirmed no
swapped button is reconfigured post-creation; the .config(state=) hits
were Text widgets). apply_ttk_styles (clam: flat accent notebook tabs,
token combobox/scrollbar) + apply_global_options (flat entries,
hairline focus ring that glows accent) wired at root init AFTER theme
load. Mapper fonts Calibri→Segoe UI (cosmetic only — structure still
untouched). Main window section frame restyled as a hairline card with
small-caps header.
**Live-tested under Xvfb** (python3-tk + xvfb installed in the dev
container): RoundButton unit behaviors (click, disable-blocks-click,
config text/state, hover, pack+grid, sizing), then full-app gauntlet
with real testkit data — main window, settings, client wizard, REVIEW
screen, the fragile MAPPER (PPTXWizard on StandIn_Report.pptx), format
popup (number AND date-only variants), query builder. All build and
render live. platform_setup (5 buttons) not live-run (needs the
add-platform file flow) — identical swap pattern to the 49 verified.
141 tests pass. Dark theme inherits everything (RoundButton reads
token values at construction).

### v16.14 — UI Refresh implementation (Phases 1, 2, 4 + wizard band)
Approved design implemented with a zero-logic-change discipline:
**Phase 1 — tokens.** config/themes.py rebuilt around the approved
palette (brand 003057 / action 0271EB / sky DEE8F5 / surface EEF2F7 /
hairline D5DEE9) with ALL legacy keys preserved — every screen that
reads self.t[...] (including the mapper) restyles automatically with
zero code changes. New additive tokens: brand/brand_fg/brand_muted,
warning, pill_ok/warn/err bg+fg. Dark theme translated to navy-dark.
ui/utils gains additive helpers: _load_logo (cached, graceful None),
brand_header (navy band: white logo + title + purely-visual step
chips 1 Files / 2 Clients / 3 Review), status_pill.
**Phase 2 — main + review + wizard.** main_window: old in-body title
strip replaced by brand_header(step=1) with Settings moved into the
band; grid row 0 left intentionally unused so NO other layout shifts.
review_view: brand_header(step=3) with client counter in-band, KPI
numbers 20pt navy with small-caps labels, hairline card borders
replacing black relief=solid, flags button on pill_warn tokens (same
widget, same behavior). client_wizard: brand_header(step=2). Font swap
Arial→Segoe UI in ui/* (review's Calibri too). Logo assets created
from repo logo.png: assets/logo_transparent.png, assets/logo_white.png.
**Mapper: untouched structurally** (fragile, per request) — it inherits
the new palette through the theme dict only; fonts left as-is.
Assignment tray + canvas badges from the proposal are DEFERRED.
**Phase 4 — Excel quick wins.** Unified Data is a real Excel Table
(UnifiedData, TableStyleMedium2 banding, own filter buttons — replaces
auto_filter, which cannot coexist), frozen header, navy tab. White
logo embedded in the Search banner (fresh layouts only; re-exports
keep their sheet). WriteResults applies native data bars to metric
columns after each search (guarded On Error, cosmetic only; the
results .Clear wipes prior conditions). DEFERRED from the proposal:
attached suggestions dropdown, removable term chips, auto KPI strip
(bigger VBA work — next round on approval).
Verified: 141 tests, smoke-import of all 9 UI modules with tkinter
stubbed, VBA structure lint, testkit rebuild incl. re-export
continuity (single Table registration, deduped rows).

### v16.38 — Search clipping ROOT CAUSE: input row never had a height
Forensics: the v16.21 row-6 move was supposed to convert the legacy
"row_dimensions[5].height" (search's old row) to row 6 — the replace
silently missed (the legacy value was 28, the pattern expected 30; no
assert). Result: the INPUT row has been at Excel's default ~15pt since
v16.21 — the true cause of every clipped-search-text field photo; the
v16.35 "fix" edited phantom lines and its verification was unreliable
(heredoc output anomalies). Now: the legacy line converted to
row 6 = 44pt with an explanatory comment, SearchCell wrap_text=True
(long searches wrap instead of running off horizontally), and the
ActiveX overlay gets MultiLine+WordWrap with EnterKeyBehavior=False
(Enter still commits). Verified on a FRESH workbook via a script file:
row6=44, row5=12, wrap on. Field note: old exports keep old layouts —
re-export to see any styling fix.

### v16.37 — Query builder field round: sidebar queries, sane preview, one output control
(1) Named queries never appeared in the sidebar: _refresh_metrics
rebuilds rows ONLY from build_simple_options(structured_metrics) — the
available_metrics registration was invisible to it. Sidebar now appends
a "Saved Queries" section from wizard.named_queries (persisted by
apply_value when a name is given); live-verified: named TABLE query
renders under the section after refresh. (2) "Numbers seem wrong":
breakdown types defaulted to SELECT-ALL, jumbling dow+dma+zip rows and
showing all-zero rows for campaign-level metrics (Contributions has no
per-level rows). Defaults to none (campaign totals); all-zero level
rows dropped; empty result falls back to campaign totals with an
explanatory note in the Total line. (3) The Output dropdown duplicated
the three Apply buttons (which override it) — removed; buttons are the
single control. 143 tests.

### v16.36 — Typed order = column order for DIMENSIONS too
Field: "Campaign, Client, Impressions" rendered Client|Campaign —
dimension columns were hardcoded Client->Campaign->Level regardless of
typed order (only metrics honored order). ParseTerms now records first
occurrence of each dimension term (campaign/client grouping words,
level types, level values) into dimOrder; RunSearch builds module
mDimSeq from typed order first, appending implied-but-untyped dims
(default campaign view, campaign-name filters) after; keyParts and
WriteResults header labels both iterate mDimSeq so aggregation keys and
columns stay consistent in any typed order. Sort grouping follows
naturally. Structure-linted; port verified 4 orderings. 143 tests.

### v16.35 — Search box text clipping (styling lane)
Committed search text rendered bottom-aligned in the cell (Excel
default) with descenders shearing on the border at laptop DPI. Input
row 6 raised 30->34 and SearchCell explicitly vertical-centered with
indent=1; the ActiveX overlay derives its height from the same row so
both rendering paths gain the headroom. Pure openpyxl.

### v16.34 — Query builder: apply row restored, named metrics (R3), taskbar-safe windows
Field: Apply as Value/Table/Chart Data row had been squeezed invisible
by the v16.31 reorder (packed AFTER the expanding preview -> zero
space) — "the way to get data out is gone". Block relocated before the
preview; live-asserted mapped and in-window at 1280x680. fit_window now
reserves the Windows taskbar zone (h <= sh-130, bottom edge <= sh-70)
so no window opens with controls behind the dock — benefits every
window. R3 completed on top of existing machinery (apply_value already
registered query metrics; engine/query_resolver already recomputes
query assignments at fill time): new "Metric name" field — a named
query registers under the human name in available_metrics, becomes the
selected metric for immediate assignment, and the sidebar refreshes
via _refresh_metrics. Auto-hash key remains the unnamed fallback.
Live end-to-end: 'Top5Networks' named, applied, registered, staged.
143 tests.

### v16.33 — Template export/import (team sharing)
engine/template_bundle.py: standardized .zip bundle (manifest.json,
template/<file>.pptx, mapping.json, images/) — export walks the mapping,
copies every referenced image into the bundle and rewrites its path to
a {{BUNDLE}}/images/ token; import extracts the pptx to templates/,
images to templates/images/<stem>/, rewrites tokens to this machine's
absolute paths, and registers the mapping. Name collision -> 
FileExistsError; the Settings handler offers replace. Foreign zips
rejected by manifest check. Settings -> Templates gains Export…/Import…
buttons with friendly dialogs + list refresh. Round-trip test suite
(export -> wipe -> import -> paths valid; overwrite; foreign-zip
rejection): 143 tests total. README team-sharing section added.
This also delivers the first half of R1: templates + images + mapping
now live together in program storage as a shareable unit.

### v16.32 — Mapper post-assign shift fixed + interactive data flags
Field (photos): shapes panel correct until an assignment renders its
long green info label -> unwrapped text widened the card -> inner frame
grew past the canvas viewport -> Assign buttons pushed out of frame.
Fixes: shapes_canvas AND metric_canvas inner frames width-locked to
their viewports (itemconfigure on Configure — review-screen pattern, no
feedback loop) + assignment labels wraplength=330. Live-verified: long
assignment label injected, frame stays locked to canvas width.
Warning tokens (warning, pill_warn_bg/fg) added to BOTH themes as
additive keys (the rollback palette lacked them — first ship attempt
KeyError'd, caught by the injected-flag live test).
Data flags upgraded from inert to interactive: campaigns implicated by
flags (tuple (campaign, metric, msg) or string mention) get an amber
outline in Campaign Detail; the flags button gains hover tooltip (first
issue + hint) and click -> detail popup listing every issue with
wrapped text, noting the amber-outline convention. 141 tests.

### v16.31 — Query builder cutoffs: root causes found + fixed
Field cutoffs reproduced/diagnosed via automated overflow detector
(/tmp/overflow_detector.py pattern: walk widget tree at Xvfb small
screens + tk scaling 1.25/1.5 + 40 long-named campaigns x 30 dma
values). Findings: (1) all windows CLEAN vs toplevel/parent bounds —
the visible "overflow" is the pivot Treeview whose COLUMNS (one per
campaign, 46+ x 110px) are internal to the widget, invisible to
geometry scans; it already h-scrolls, columns now stretch=False.
(2) REAL bug: refresh button + export options packed LAST -> squeezed
off the bottom on short screens (the half-cut "Refresh Preview" in the
field photo). Fixed by packing bottom controls FIRST side="bottom"
(review-screen fix class); the scrollable preview absorbs deficit.
Verified at 1280x680 + 1.25 scaling with fat data after a real refresh:
button on-screen, 40+ column pivot scrolls. Export-columns listbox
showing campaign names is CORRECT (pivot columns ARE campaigns).
R3 CONFIRMED by John: named queries recompute on every re-fill.

### ROADMAP (clarified with John, pre-team-testing) — build order
R1. Program-folder storage: selecting a template COPIES it to
    storage/templates/<TemplateName>/ (template + browsed images +
    mapping.json live together); exports default to output/. Old
    absolute image paths keep working and self-heal into storage on
    next save. Ships first — everything else builds on this layout.
R2. Learned suggestions: history file records (placeholder text,
    shape name) -> metric per assignment; keyed primarily on
    PLACEHOLDER TEXT so learning transfers across templates. 3+
    consecutive identical pairings -> suggestion strip under the shape
    card ("Usually: Client Name" + Quick Fill) + "Apply all
    suggestions". Never auto-applies.
R3. Named queries as metrics: Query Builder gains a name field; saved
    query appears in the sidebar metrics list, assignable to shapes;
    persisted in the mapping and RECOMPUTED on each re-fill with fresh
    data (pending John's confirm).
R4. Dynamic chart fill (comparison charts): introspect the chart's own
    structure (series count, category orientation, series names)
    before writing; map assigned data into THAT layout instead of
    assuming categories-down-column-A single-series. Field evidence:
    normal Top-5 bars fill fine; March-vs-April comparison charts use
    series-across-columns and currently mis-place data. Largest item;
    touches the fragile fill path — proposal-first, ships alone.

### v16.30 — Portable mode (vendor folder) + team-testing handoff
Additive vendor shim in main.py: if ./vendor exists it is preferred on
sys.path (incl. win32/Pythonwin subpaths + pywin32_system32 DLL dir for
COM); absent, a strict no-op — normal startup cannot be affected.
make_portable.bat populates vendor/ via pip --target. README section
added. FIELD NOTEBOOK LOGGED for next rounds (clarifications pending):
(1) 3x-repeated metric->shape assignment triggers auto-suggest;
(2) comparison charts layout issue (detail needed); (3) Advanced Query
gets a name + output saved as a named metric for templates; (4) browsed
images/pptx stored in program storage (temp -> template-named folder).

### v16.29 — Campaign rows only when explicitly requested (spec change)
Original grammar (per the first spec) made level-VALUE searches show
campaign rows ("28167" -> campaigns under it). New approved rule: a
value search shows THAT LEVEL as the rows ("10603, impressions" ->
row 10603), aggregated across campaigns; the Campaign column appears
ONLY when explicitly typed — the "campaign"/"client" grouping words or
an actual campaign name (a typed name both filters AND shows the
column). Empty search keeps the campaign-totals overview (the no-query
state). Implementation: RunSearch derives levelType from the first
level value's prefix when no type term was given; useCamp requires
campGroup Or camps.Count>0. Port-verified: value-only -> level rows,
+"campaign" -> compound, +name -> filtered compound. 141 tests.

### v16.28 — Mapper shapes panel guaranteed its width (assign access)
Field: window fit the screen (v16.24) but the RIGHT shapes panel was
clipped — grid gave sidebar + preview FIXED widths (preview hardcoded
680) and shapes only got the remainder (~270px at a clamped 1240px
window), cutting the cards and their assign targets. Fix: the preview
frame's LOCKED size is now computed from the screen (win_w - 720 for
sidebar+shapes+pads, capped 680x520 so big monitors are unchanged,
floored 400x320) — the shapes panel is guaranteed ~430px everywhere.
slide_view's refresh now locks the image to the preview label's actual
size (propagate stays off — no feedback loop), completing the
container-locked behavior. Live at 1280x800: shapes panel >=380px,
right edge inside the window, first card fits the panel. 141 tests.

### v16.27 — Review screen: accidental sidebar fixed, content width-bound
The "Bottom action bar" packed AFTER the side-left content canvas, so
pack geometry shoved it into the top-right cavity as an accidental
sidebar overlapping/clipping the KPI cards (field photo: 4th card cut
behind Generate Report). Fixes: (1) action bar packs side=bottom BEFORE
the canvas — a true full-width bottom bar; (2) the scrollable inner
frame is width-bound to the canvas (itemconfigure on canvas Configure —
sets item width only, so no resize feedback loop is possible), letting
KPI rows and campaign cards wrap to the visible width on any screen.
Live-verified at 1280x800 with screen coordinates: canvas spans the
window, Generate Report sits below the content. 141 tests.

### v16.26 — Completion Rate in the Excel search (approved rule, VBA round)
Single-purpose VBA round re-applying the approved 100*Σcompletions/
Σimpressions to the search dashboard (the app Review has had it since
v16.19; the two surfaces now agree). _Config marks Completion Rate/
Percent agg="rate"; modSearch accumulates Impressions + PickNumerator
(Contributions > 100% Completions) via Chr(4) pseudo-keys riding the
existing source collapse even when not displayed; WriteResults renders
100*num/imp as 0.00% per row with graceful fallback to the averaged
value when denominators are absent; the KPI strip computes rate cards
as grand sum/sum (never summed percentages). Structure-linted; math
port verified per-campaign and grand against a generated workbook;
142-anchor edits all asserted; 141 tests.

### v16.25 — ALL windows responsive: fit_window sweep
Audit found 10 fixed geometry calls (main 800x500 x2, settings 560x520,
rename 350x130, wizard 750x650, platform setup 750x700, platform-pick
300x280, review flags 600x420, query builder 950x700, format popup
360x300/540). All routed through new ui.utils.fit_window: desired size
clamped to the actual screen (margins 40/90) and centered — big
monitors unchanged, laptops always fit. Live-verified at a harsh
1024x700 Xvfb screen: main, settings, wizard, mapper, query builder,
and format popup all fully on-screen.

### v16.24 — Mapper window clamps to the screen (thumbnail root cause 2)
The v16.22 fix clamped the PREVIEW to the window, but the window itself
was fixed at 1400x800 — wider than DPI-scaled laptop screens when not
zoomed, so it hung off-screen and clipped the preview with it. Geometry
now clamps to winfo_screen size (margins 40/90) and centers; the
preview bounds already derive from the window, so the whole UI scales
with screen size as requested. Live-tested under Xvfb at a forced
1280x800 screen: window fully on-screen and centered. Desk monitors
(>=1440 wide) get the same 1400x800 as before.

### v16.23 — Suggestion chips complete the LIVE draft (field bug)
ApplySuggestion read SearchCell — the previously COMMITTED search — so
clicking a chip mid-typing resurrected the old query and appended
(field: click Impressions -> a ton of terms appear). Now reads the
overlay textbox draft first (SearchCell fallback when no overlay), so
a chip only completes the last fragment of what is being typed — the
text-based-slicer behavior. Also: overlay font 12->10 + Segoe UI and
cell font 12->11 (12pt clipped descenders in the 30px search box).

### v16.22 — Mapper preview: adaptive-capped bounds (audit + fix)
AUDIT (field report of the classic thumbnail bug): mapper/ and
ui/utils.py byte-diffed against field-tested 1.13 — IDENTICAL; no
change from v16.19-16.21 touched the mapper. Root cause: the historical
fix used HARD-CODED 660x490 preview bounds tuned for the desk monitor;
on a windowed/DPI-scaled laptop the fixed pane exceeds the window and
clips — latent in 1.13, not a regression. Fix (contained to
_refresh_preview): bounds computed from the toplevel size at refresh
time, capped at 660x490 (large screens unchanged), floored 320x240, NO
Configure binding so the original growth feedback loop cannot return.
Mapper live-built under Xvfb incl. the changed path; 141 tests.

### v16.21 — Dashboard breathing room (field round 2)
Excel: airy grid per feedback — spacer COLUMNS G/I/L between every
element and spacer ROWS 4/7/10; inputs move to row 6 (labels row 5),
hint gets its OWN row 8, suggestions row 9 with chips on ALTERNATING
columns (B/D/F/H/J/M, cap 6 — no more clustered run-together chips),
Results row 11 + copy chip D11, results start row 12. Button column
widened to 14 ("Sear" clipping fixed). Frame ends ~col N (~1050px),
logo at L1. Lockstep: python constants/defined names, injector overlay
B6:F6 + sheet-code addresses, modSearch SUGGEST_ROW=9/RESULTS_ROW=12 +
copy-chip cell + chip column array. App (safe lane): review KPI cards
wrap at 4 per row so they never run off-screen. KPI strip unchanged
(field-approved).

### v16.20 — Single-frame Excel dashboard + safe app branding
Excel (aggressive, per approval): compact grid fits columns A–K at 100%
zoom for zero-scroll users — widths 16/11/12/4, search B5:F5, button
G5, dates H5/J5 + chips I5/K5 (lockstep: constants, defined names via
constants, injector overlay B5:F5, sheet-code addresses). Banner/accent/
rule span A–K only; hint moved under the search box; logo 26px at I1.
VBA WriteResults gains the deck's AUTO KPI STRIP (merged 2-col cards:
15pt totals + small-caps labels for up to 4 visible metrics, sums/avg
per _Config) and guarded native data bars on metric columns. App
(cautious): theme palette values only — 16 keys, names unchanged, every
screen incl. mapper restyles through self.t reads — plus Segoe UI in
ui/* (mapper untouched). Verified: 141 tests, VBA lint, frame width
~123 units (~890px), app live-built under Xvfb.

### v16.19 — ROLLBACK to 1.13 baseline + safe-only re-applies
App UI surgery (v16.14–16.18: tokens, RoundButton, brand bands, mapper
fonts) fully reverted after field bugs — ui/, mapper/, themes.py, and
the VBA search module are byte-identical to field-tested 1.13.
Re-applied ONLY zero-risk items on that base:
- Completion Rate = 100*Σcompletions/Σimpressions in engine/kpi.py
  (approved rule; python-only — Excel VBA stays 1.13, so its search
  still shows the averaged rate until a careful VBA round).
- Excel styling (pure openpyxl, cannot affect logic): Search sheet as
  a dashboard (row/col headers hidden, white canvas, full-width navy
  banner + accent, white logo in banner), Unified Data as a real Excel
  Table + frozen header + navy tab.
- Two field-bug fixes to code that predates the surgery: ActiveX
  overlay positioned via properties after Add (was landing at 0,0 —
  the "white bar on A1"), and xlsx-after-xlsm delete retried 6x with
  writer-side stale-pair self-heal (was leaving dual files).
Deferred until requested: all app-UI restyling, VBA rate math, deck's
remaining Excel features (dropdown/chips/KPI strip).

### v16.13b — Proposal expanded: popups + Excel workbook
Deck now 14 slides: added Advanced Query Builder, Format popups (number
+ date-only) with the shared calendar, Platform Setup with the
rate-suspicion warning surfaced at setup time, shared dialog language
(verb-button confirms, icon toasts) — and two Excel slides: Search
dashboard v2 (attached suggestions dropdown, removable term chips, auto
KPI strip for the current search, native data bars) and Unified Data
polish (real Excel Table, frozen header, brand tab colors, "⚙ Columns"
chip opening _Config). Rollout is now 4 phases, Excel as Phase 4.
Still proposal-only — nothing implemented.

### v16.13 — Documentation sweep + UI refresh proposal
- README, docs/user_guide.md, docs/ARCHITECTURE.md, docs/API_REFERENCE.md
  updated to the current export model: Search dashboard + Unified Data,
  write_to_excel's (path, vba_error) contract, retired settings keys,
  Trust Center setup, and a user-facing search-grammar reference table.
- UI_Refresh_Proposal.pptx/.pdf delivered for review (NOT yet
  implemented): design-token system on brand colors (navy 003057 /
  blue 0271EB / sky DEE8F5 / surface EEF2F7), component library,
  per-screen mockups (main, review, mapper, wizard, settings), honest
  tkinter constraints, 3-phase rollout. Implementation starts only
  after approval.

### v16.12 — One-click Copy results
"📋 Copy results" chip beside the Results label (C9 — above the results
area so table clears never remove it). WriteResults records the exact
extent of the displayed table (header row / last row / last col);
clicking the chip copies that range — header + rows + formatting,
excluding the "Ignored:" line — to the clipboard, flashes
"✓ Copied N rows" for 2s (Application.OnTime restore), or "Run a search
first" when there's nothing to copy. Note: the chip emoji needs a VBA
surrogate pair (U+1F4CB is above the BMP — ChrW$(&HD83D) & ChrW$(&HDCCB)).

### v16.11 — Search grammar v2: grouping, OR values, perf, integrity
From the second field round (VBA modules now live in engine/vba_src/):
- **"campaign"/"client" are grouping terms** (row_group in _SearchIndex):
  "Campaign, Zip Code, Impressions" now yields compound rows with
  Campaign + Zip label columns instead of "Ignored: Campaign". Combine
  freely — client, campaign, and a level type can all be row dimensions.
- **Multiple values OR together**: "28167, 28105, impressions" filters
  to both zips. Mixing level TYPES in values (Monday + CNN) keeps the
  first type and flags the rest — no parentheses/+ needed; commas AND
  across kinds, repeat values to OR within one.
- **Duplicate metric columns deduped**: "Contributions, Completions"
  both alias to Contributions and now produce ONE column.
- **Contains-matching added** (exact -> prefix -> contains); in the
  contains pass level VALUES drop to lowest priority so "work" finds
  the network TYPE, not "ACC Network".
- **Performance**: _SearchIndex/_Config were read cell-by-cell over COM
  4 passes per term and per keystroke — now bulk-loaded into arrays
  once (InvalidateCaches available). This was the perceived slowness.
- **Display**: non-whole additive values render 2dp (the "Contributions
  all 0" was 0.03% rate values truncated by the integer format — the
  Architect rate-alias issue, see v16.10).
- **UI**: labels above inputs, wider search box (B5:G5), button H5,
  dates I5/K5 + chips J5/L5, suggestions row 7, results row 10 — all
  addresses changed in lockstep (layout constants, modSearch.bas,
  sheet code, injector overlay, defined names).
- **Data integrity PROVEN**: source-parse totals vs Unified Data sheet
  totals match exactly for every metric (7/7 on testkit).
- Grammar scenarios (including the exact field-reported failures)
  verified via a Python port of the VBA against a generated workbook.

### v16.10 — Search Dashboard field fixes (first live VBA run)
Field test confirmed injection, search grammar, chips, and xlsm output
all work. Fixes from the run:
- **Calendar compile error**: the composed frmCalendar source placed
  `Private mHandlers As Collection` after procedures — VBA requires all
  module-level declarations first ("Only comments may appear after End
  Sub"). Declarations are now hoisted to the top of the composed source.
- **True as-you-type suggestions**: cell typing only fires events on
  Enter, so the injector now overlays an ActiveX TextBox ("txtSearch")
  on the search cell; its Change event updates suggestion chips per
  KEYSTROKE, Enter (or the Search button) commits to SearchCell which
  triggers the run. Fully guarded: if the control can't be created,
  cell input + Enter still works exactly as before.
- **Spectrum Reach restyle**: brand navy (003057) banner with tag line,
  bright-blue (0271EB) accent rule/button/tab, thick blue Results rule —
  and the cut-off search box fixed: openpyxl styles only the top-left
  cell of a merged range, so `_style_range` now paints borders/fills on
  EVERY underlying cell. All VBA-relied addresses unchanged (B4/H4/J4/
  row 6/row 9).
- Alias coverage question answered from the dictionary: 19 metric
  families and 11 level types carry aliases; unmatched raw column names
  remain searchable verbatim. NOTE from field data: Architect's
  "Contribution" column carries PERCENT values (0.03%), alias-folded
  into additive Contributions — the confirmed root of the
  "Completions: 1.00" mystery. Fix belongs in the Architect platform
  config (remap that column to Completion Rate), not the global alias.

### v16.9 — Interactive Search Dashboard (Excel), export simplification
The unified workbook is redesigned around a single smart-search sheet.
NOTE: this feature was found partially built in the workspace (a prior
session was cut off mid-task); the existing modules were audited line by
line, two real VBA bugs were fixed, and the remainder completed.

**Workbook** = "Search" + "Unified Data" (+ hidden "_SearchIndex",
"_Config"). Client Report and the old Campaign Dashboard sheet are gone,
as are the export checkboxes and the Dashboard settings tab (replaced by
the _Config sheet: unhide it to change default metric columns/order/
aggregation — the Alteryx-style column control).

**Search grammar** (one box, comma-separated, broad-to-specific):
level value ("28167", "CNN") -> rows = campaigns under it; level type
("zip", "network") -> rows = that type's values; metric terms restrict
AND ORDER the columns (typed order = column order — that is the quick
column-reorder mechanism); campaign terms filter. Aliases resolve via
_SearchIndex (built from the metric dictionary + data); level values
are searchable WITHOUT their "prefix:" and date terms drop Excel's
" 00:00:00" noise while canonical filters stay verbatim. Totals view
uses best-source-per-campaign, matching engine.kpi.

**VBA engine** (engine/excel_vba.py) injected at export via Excel COM,
saved as .xlsm: type-ahead suggestion chips (click to complete),
Search-button chip, live re-run on change, runtime-built calendar
UserForm for the date filters, sum/avg per _Config. Requires Excel's
"Trust access to the VBA project object model" once; on ANY failure the
plain .xlsx ships with data intact and the UI explains what to enable
(write_to_excel returns (path, vba_error)). Re-exports append into the
existing .xlsm, dedupe rows, and never touch the Search sheet (its VBA
code-behind must survive).

**Fixed during audit**: calendar month-nav buttons were runtime-created
but wired to compile-time Click handlers (which never fire for
Controls.Add controls) — routed through the WithEvents relay class;
unsafe control-collection mutation during For Each replaced with a
drain loop. Off-Windows fallback, index/config builders, re-export
continuity, and dedup covered by tests/test_search_dashboard.py
(140 total). VBA block structure machine-checked (continuation-aware);
runtime behavior still needs a desk test (COM cannot run here).

### v16.8 — Dead-code sweep (pyflakes + vulture verified clean)
Holdover audit after the heavy mapper churn. Removed, with caller
verification on every item:
- `engine/pptx_live.py`: `show_window`, `hide_window`, `get_slide_count`,
  `get_shape_text` — monolith-era API, zero callers.
- `engine/pptx_mapper.py`: `get_internal_template_path` and stale
  re-export imports of pptx_formats helpers nothing imported through it.
  The live re-exports (`fill_template`, `get_available_metrics`,
  `SPECIAL_METRICS`) are kept and now explicitly marked `# noqa: F401`.
- `engine/query_resolver.get_available_metrics_list`,
  `engine/data_pipeline.get_all_campaign_names` (the client wizard uses
  its own richer platform-attributed collection — the plain helper was an
  orphaned simpler variant), `engine/errors.PipelineError`,
  `config/paths.BASE_DIR`.
- `ui/utils._bind_mousewheel` no-op shim plus all its imports/calls
  (review_view, mapper_window) — `enable_mousewheel` binds app-wide and
  covers Toplevels.
- `write_to_excel` legacy `agg_map`/`table_defs` parameters (+ call site).
- Dead locals/imports: parser `metrics = {}` leftovers, unused
  `load_dictionary`/`ttk` imports, `level_key` loop var, sidebar `t`,
  settings `MAPPINGS_DIR`/`LOG_FILE`, `fmt_var_unused` parameter on
  `_assign_to_shape`.
- **Holdover hazard fixed** (`mapper/slide_view.py`): pending chart/table
  data was gated on `hasattr` — once set, STALE data could re-apply to
  later assignments forever. Now value-checked and consumed after use.
Tests updated for the removals; suite at 132 passing; vulture reports
zero dead functions/methods; pyflakes shows only the two documented
re-export blocks.

### v16.7 — Reformat-in-place; date-only format popup
- **Assignment lists were wiped by re-assigning** (`mapper/slide_view.py`):
  the mapping write only appended when the user had highlighted replace
  text. Re-assigning a metric with nothing highlighted (the natural move
  after changing its format, since the placeholder text is already
  replaced) hit the fallback branch and OVERWROTE the shape's whole
  assignment list — then the full replace landed on the first paragraph
  ("assigned to the wrong text"). Now: re-assigning a metric already on
  the shape updates that assignment in place; a genuine full-replace
  onto a shape with assignments asks for confirmation; appends work with
  or without prior structure (legacy single-style entries normalized).
- **Reformatting never requires re-assigning**: the format popup's OK now
  propagates new details into every existing assignment of that metric
  (`_propagate_format_details`) and re-renders the live slide in place.
- **last_rendered idempotency** (`engine/pptx_live.py`,
  `mapper/slide_view.py`): every assignment records what it rendered;
  live updates take a candidate list [last render, original placeholder]
  and replace the first one found — so chained format changes replace
  their own previous output, verified headlessly:
  "May 1st, 2026" -> "05/01/2026" -> "May 2026" with two assignments on
  the shape surviving untouched throughout.
- **Date-only format popup** (`mapper/format_popup.py`,
  `mapper/sidebar.py`): right-clicking Date Range / Start Date / End
  Date opens a compact date-style chooser (dropdown + custom strftime +
  preview) — no number-formatting noise, per the field suggestion.

### v16.6 — Shape-identity image fills, dynamic KPI cards, date formats
- **Wrong-object assignments after image swaps** (`engine/pptx_live.py`,
  `mapper/slide_view.py`): `replace_shape_with_image` deleted the shape
  and appended the picture at the END of the Shapes collection, shifting
  every later shape's index — subsequent assignments (text, custom text,
  images) silently hit the wrong objects. This was the root of BOTH the
  image bug and "custom text says assigned but does nothing." Now:
  regular shapes get a picture FILL (identity preserved; label text is
  snapshotted and blanked); real pictures are swapped and walked back to
  the deleted shape's exact z-order slot. `update_shape_text` also takes
  `expected_name` and logs "SHAPE INDEX DRIFT" loudly if the shape at an
  index no longer matches. Known limit: Clear All on a picture-filled
  shape restores text/geometry but not the original fill.
- **Browsed images are reusable Quick Fill entries** (`mapper/sidebar.py`):
  each browsed picture registers as its own sidebar entry
  ("🖼 Image: name.png") under Browse Image — re-selectable without
  re-browsing, no shared pending-image state.
- **KPI screen shows ALL metrics** (`ui/review_view.py`): the hardcoded
  [Impressions, Clicks, Completions, Cost] whitelist became a preferred
  ORDERING; every computed metric renders (5 cards per row), and each
  campaign's summary line shows its own top metrics.
- **Rate-suspicion data flag** (`engine/kpi.py`): an additive metric
  whose row values are all ≤1.5 across ≥5 rows (e.g. "Completions:
  1.00") is flagged as a probable mis-aliased rate/percentage.
- **Excel-style date formatting** (`engine/pptx_formats.py`,
  `mapper/format_popup.py`): Format popup gains a Date category with a
  style dropdown (May 1st 2026 / 05/01/2026 / May 2026 / ISO / ... )
  plus custom strftime. `format_date_with_style` handles single dates
  and ranges; date details travel with assignments through the live
  preview and both fill paths. (134 tests.)

### v16.5 — Metric aggregation types (sum vs average)
Field test showed "Total Frequency = 159.74" — 63 daily frequencies of
~2.5 summed into a meaningless number, with the raw float repr
(159.74153592699997) leaking into the format preview. Root cause: every
metric was treated as ADDITIVE. Changes:
- `metric_dictionary.json` gains `metric_aggregation` declaring ratio
  metrics (Frequency, CTR, CPM, CPC, Completion Percent, Completion
  Rate, Impression Share) as `avg`; everything else defaults to sum.
  Add new ratio metrics there — no code change needed.
- `engine/kpi.py` + `engine/metrics_catalog.py`: ratio metrics average
  per campaign, then across campaigns. Catalog keys become honest:
  "Avg Frequency" (not "Total Frequency"), "Date Avg: Frequency", and
  per-item breakdown values are means. Additive behavior unchanged.
- Sidebar labels/queries (`engine/query_resolver.build_simple_options`)
  say "Avg X" with `agg: avg` for ratio metrics.
- Float repr leak fixed: default text formatting renders non-whole
  floats at 2dp everywhere (popup preview, sidebar, insertion).
- CAVEAT: Reach remains additive because true deduplicated reach cannot
  be derived from exports — summed Reach over days double-counts
  overlapping audiences. Treat "Total Reach" as an upper bound.
- Regression tests added (132 total).

### v16.4 tweak — auto-widen returns, done right (`engine/pptx_live.py`)
Field test confirmed the COM live preview works, but long values wrapped
in narrow stat boxes ("3,027,614" split across two lines). The original
auto-widen concept was correct; its implementation was the bug (failed
measurements permanently disabled WordWrap). New `_widen_to_fit`:
measures the one-line width via a temporary WordWrap toggle + BoundWidth,
grows the shape SYMMETRICALLY (left shifts half the delta so centered
stat blocks stay centered), clamps to slide bounds, and on measurement
failure falls back to shrink-text-to-fit — never to disabling WordWrap.
Snapshots now capture left/width/autosize so Clear All restores geometry
and the template's native autosize setting too.

### Found by KPI stress testing (multi-platform ground truth)
- **PowerPoint totals silently dropped campaigns** (`engine/metrics_catalog.py`):
  "Total X" metrics took the max of GLOBALLY-summed sources, while the
  review screen takes the best source PER CAMPAIGN and sums. Any campaign
  whose best source differed from the group's (no summary sheet,
  cross-platform clients) vanished from the deck total — verified with
  ground truth: review 1,505 vs PowerPoint 1,005. metrics_catalog now
  uses the same per-campaign-best rule; deck and review always agree.
  Verified exact at scale (58k rows, 40 campaigns, 4 sources) and across
  3 platforms (xlsx + CSV with alias columns + HTML) in one client.
- **HTML files contributed zero data** (`parsers/html_parser.py`): the
  HTML parser detected tables but never classified them — no
  level_data/campaign_metrics keys at all, crashing the pipeline (or
  silently contributing nothing) unless a platform config rebuilt them
  with exactly matching sheet names. It now self-classifies via the
  shared dictionary exactly like the CSV/Excel parsers.
- Regression tests added: `tests/test_kpi_consistency.py` (130 total).

### Known behavior surfaced by testing (not changed)
- Breakdown sheets WITHOUT a campaign column (e.g. a Devices export with
  no "Campaign Name") are silently dropped by client campaign filtering
  — their rows carry an empty campaign that never matches an assigned
  campaign list. If device breakdowns are missing from client reports,
  this is why. Needs a product decision: keep, warn, or pass through.

All 126 tests pass.

---

## v16.2 — Beta 1.2 Field-Test Fixes (July 2026)

### Bug Fixes
- **`filedialog` NameError on Auto-Fill** (`ui/review_view.py`): the
  template selector's save dialog referenced `filedialog` without
  importing it. Import added.
- **Files could not be re-added after a run** (`ui/main_window.py`):
  `_build_ui()` reset `file_widgets` but left `selected_files`
  populated, so the dedup check silently swallowed re-added files after
  Finish / Back-to-Import (and the stale list desynced from the widget
  rows). `_build_ui()` now resets `selected_files` too — bulk add via
  `askopenfilenames` was always present and works again once state is
  in sync.
- **Multi-line text boxes mangled on assignment** (`engine/pptx_live.py`):
  in PowerPoint COM, `Paragraph.Text` includes the trailing `\r`
  paragraph mark; assigning without it merged the paragraph into the
  next one. That is what made "AIRINGS"-style sublabels vanish, big-font
  runs swallow the whole box, and long lines wrap. Paragraph marks are
  now preserved (`_set_paragraph_text`). Overflow is handled with
  shrink-text-to-fit (`TextFrame2.AutoSize = msoAutoSizeTextToFitShape`)
  only when the replacement is longer than the original.
- **Clear All was unreliable** (`engine/pptx_live.py`,
  `mapper/slide_view.py`): Clear used a single global
  `ExecuteMso("Undo")`, but assignments are re-applied on every slide
  navigation, piling up COM operations that one undo cannot unwind (and
  it could undo the wrong operation). The live preview now snapshots a
  shape's per-paragraph text before its first modification and
  `restore_shape_text()` puts it back exactly, per paragraph, preserving
  each paragraph's formatting.
- **Pictures invisible to the mapper** (`engine/pptx_mapper.py`): the
  scanner only surfaced shapes with text frames, tables, or charts —
  plain pictures (e.g. the zone map) never appeared. Both scan paths
  (python-pptx and COM) now detect picture shapes
  (`MSO_SHAPE_TYPE.PICTURE`/`LINKED_PICTURE`, `msoPicture`/
  `msoLinkedPicture`) and list them as `[Image: name]`.
- **Format popup preview showed 0** (`mapper/sidebar.py`,
  `mapper/format_popup.py`): query-backed sidebar entries (KPI totals
  etc.) resolved their values on the fly without caching, so the
  right-click format preview looked them up in `available_metrics` and
  fell back to 0. Resolved values are now cached, and a truly missing
  value shows "(no value yet)" instead of a fake 0.

### New Features
- **Query builder export controls** (`mapper/query_builder.py`,
  `engine/pptx_live.py`): "Columns to export" multi-select (refreshed
  from the live pivot, label column included), a "table already has
  headers" toggle that writes data below row 1 without overwriting the
  template's headings, and an optional comma-separated custom-header
  override for table/chart output.
- **Settings → Templates → Edit Mapping** (`ui/settings_window.py`,
  `mapper/mapper_window.py`): opens the full mapper on a saved template
  with its mapping preloaded so assignments can be fixed or added
  without running data. `PPTXWizard` accepts `template_path=` to skip
  the file dialog; empty client data is handled by `metrics_catalog`.
- **Static slide-1 thumbnails** (`engine/pptx_thumbs.py`): templates
  now show a real screenshot of slide 1 (COM PNG export, cached under
  `templates/thumbs/` and regenerated when the template file changes)
  in both the Settings Templates tab and the report template selector,
  with the old text summary as fallback. Rename/Delete keep the cache
  in sync.

### Policy
- The 400-line-per-file hard limit from the v16.0 refactor is retired.
  Modules should stay focused (split when a file covers more than one
  concern), but there is no numeric cap.

All 126 tests pass.

---

## v16.1 — Beta 1.1 Bug Fix + Feature Release (July 2026)

Field-tested fixes and feature additions from beta 1.0:

### Bug Fixes
- **Thumbnail growth on metric assignment** (`mapper/mapper_window.py`,
  `mapper/slide_view.py`): Preview frame locked to fixed dimensions with
  propagation disabled; center column weight set to 0; image scaling
  changed to aspect-ratio-preserving `thumbnail()`.
- **Calendar arrow clicks closing the popup** (`ui/main_window.py`):
  `<FocusOut>` handler now checks `winfo_toplevel()` of the newly-focused
  widget — only closes when focus truly leaves the popup.
- **Mousewheel scrolling not universal** (`ui/utils.py`,
  `ui/main_window.py`): Replaced per-canvas enter/leave bindings with a
  single `enable_mousewheel(root)` call that routes scroll to the
  scrollable widget under the cursor. Added Linux Button-4/5 support.
- **Finish button TclError** (`ui/review_view.py`, `ui/main_window.py`):
  `_show_final_summary()` now calls `_build_ui()` first to recreate the
  destroyed `status_label`. Window state restores to 800×500.
- **Template auto-fill producing nothing** (`ui/review_view.py`): Added
  save-as file dialog (was silently writing to a path that could fail),
  added missing `__client_name__`/`__start_date__`/`__end_date__` meta-keys,
  added close-out process with option to open the generated file.
- **Browse popup lowering window priority** (`mapper/sidebar.py`,
  `mapper/mapper_window.py`): Added `parent=self.window` to all
  `filedialog` calls in the mapper so file browsers don't drop the
  mapper window behind the main window.
- **Text wrapping / assign-clear-assign cycle** (`engine/pptx_live.py`):
  Removed aggressive auto-widen logic that broke shape dimensions and
  permanently disabled WordWrap on fallback. Text now wraps naturally
  within the original shape bounds.
- **Slide preview too small** (`mapper/mapper_window.py`,
  `mapper/slide_view.py`): Preview frame increased to 680×520, thumbnail
  max to 660×490 — nearly 50% more area for visual clarity.

### New Features
- **Settings → Templates tab**: List all saved templates with
  mapped/unmapped status, slide-1 text preview panel, rename (with
  mapping file rename), and delete (with mapping file cleanup).
- **Settings → Debug tab**: Live log viewer (last 200 lines, Consolas
  green-on-black terminal style), Refresh/Open Logs Folder/Clear Logs
  buttons, log file path display.
- **Template selector thumbnails** (`ui/review_view.py`): Template
  selector now shows a slide-1 content preview panel when selecting a
  template, wider layout (600×420).
- **Dashboard Metrics tab** renamed to "Dashboard" for cleaner tab bar.

All 126 tests pass.

---

## v16.0 — Production Readiness Release (July 2026)

Post-15.5 engineering pass; see git history for the step-by-step record:
- **Architecture:** monolith split into `config/`, `ui/`, `mapper/`,
  `engine/` per ARCHITECTURE.md; modules kept focused and reasonably sized; `main.py`
  entry point; ~380 lines of dead code removed (ReviewWindow, wizard
  date step).
- **Reliability:** structured `ParserError`s with user-facing messages;
  zero bare excepts; rotating file logging (`logs/`); global crash hooks
  (sys.excepthook + Tk report_callback_exception); per-file import
  failures no longer abort Quick Run.
- **Bug fixes:** pandas float noise cleaned in `_format_value`
  (1.0000000000000009 → 1); whole numbers display without decimals;
  `short_month` date style renders; theme switch no longer duplicates
  the file list; locked/corrupt output files produce clear messages.
- **Tests:** 126-test pytest suite for the data pipeline
  (`python -m pytest tests -q`).
- **Packaging:** hardened `launch.bat` (Python detection, 3.10+ gate,
  install-only-if-missing); pinned `requirements.txt`; PyInstaller
  onedir spec (`ingestion_engine.spec`) as fallback.
- **Docs:** `docs/` — User Guide (md+PDF), Technical Documentation
  (md+PDF, supersedes API_REFERENCE.md), July 14 demo script, manual
  verification checklist.

Known issues #2 (dead code), #3 (COM error handling) below are resolved;
#4–#8 remain as post-demo work.

---

## Project Summary
**Name:** Spectrum Reach Reporting Ingestion Engine
**Version:** 16.0 (UI label v2.0)
**Author:** John (Spectrum Reach Internship — Client Development & Performance Team)
**Purpose:** Automate campaign reporting by ingesting data from multiple platforms, generating Excel dashboards, and filling PowerPoint templates.

## Current Version History (v10–v15.5)

### Data Pipeline (v10–v13)
- v10: Universal metric dictionary module (`parsers/dictionary.py`)
- v11: Unified platform setup from sample files
- v12: Wizard-based client assignment, per-client folder export
- v12.2–12.5: Three-sheet Excel (Client Report, Dashboard, Unified Data), export options
- v12.6: Review popup with KPI cards
- v12.7: PowerPoint template mapper (click-to-assign)
- v12.8–12.9: KPI calculation fix, autofilters
- v13.0: **Pandas refactor** — entire data pipeline rebuilt with pandas
- v13.1: COM-based live PowerPoint preview
- v13.2: Mousewheel + window focus fixes
- v13.3–13.5: KPI totals using pandas groupby (max across all sources)

### PowerPoint Automation (v13.6–v14)
- v13.6: Format-preserving text replacement (run-level)
- v13.7: Date auto-format + undo via COM
- v13.8–13.9: Query resolver system (simple + advanced)
- v14.0: Date format dictionary (metric_dictionary.json)
- v14.1: Expanded KPI aliases, COM undo fix
- v14.2: Smart text case/spacing matching (CamelCase split, case matching)
- v14.3: Visual pivot table in query builder, chart/table data via COM
- v14.4: Re-apply assignments on slide navigation
- v14.5: COM fallback scan, date format fixes
- v14.6: Multi-assignment per shape
- v14.7: Start Month, End Month, Year metrics, custom text field
- v14.8: Format popup removed from shapes, right-click format on sidebar
- v14.9: COM text replacement using VBA approach (Paragraphs(1).Text)

### UI Consolidation & Polish (v15)
- v15.0: Template/image portability (internal copies, relative paths)
- v15.1: Multi-select query builder (campaigns, breakdowns, values)
- v15.2: Review merged into main window, embedded slide thumbnail preview
- v15.3: Date range on main screen, simplified Quick Run, default output folder
- v15.4: Settings Export tab, calendar buttons restored
- v15.5: Maximized windows, calendar dropdown, format popup centered, number formatting fixes, floating point precision fix

## Unified Data Schema (9 columns)
```
client | campaign | campaign_type | source | metric_level | metric_name | metric_value | start_date | end_date
```

## Key Technical Decisions

### Data Pipeline
- **Pandas throughout** — groupby().sum() for deduplication/aggregation, pivot_table() for wide format
- **Metric dictionary** (`metric_dictionary.json`) — aliases, level definitions, date formats, context columns, skip columns
- **Platform configs** — saved per-platform in settings.json, column roles: metric, campaign_id, level:type, skip
- **KPI calculation** — sum by source type per campaign per metric, take MAX across sources (handles missing campaign summary sheets)

### PowerPoint COM Automation
- **python-pptx** for file-based scanning and filling
- **win32com (pywin32)** for live preview, chart/table updates
- **Text replacement** uses VBA approach: `Paragraphs(1).Text = value` (no Select, no Characters for full replace)
- **Partial replacement** uses `Characters(idx, len).Text = new_text`
- **AutoSize = 0** (ppAutoSizeNone) — lets PowerPoint handle wrapping natively within shape bounds
- **Slide export** as PNG for embedded thumbnail preview
- **PowerPoint runs minimized** — no visible window, only thumbnail shown

### Template Mapping System
- Mappings saved by template filename in `/mappings/pptx_{filename}.json`
- Templates copied to internal `/templates/` folder
- Images copied to `/templates/images/`
- Relative paths stored for portability
- Multi-assignment per shape (list of assignments with individual replace_text)
- Query-based assignments are reusable across clients

### Date Format System
- Date formats defined in `metric_dictionary.json` under `date_formats`
- Auto-detects format from existing PowerPoint text
- Converts data dates to match (long_ordinal, us_slash, iso, month_only, month_year, etc.)
- Handles brackets, misspelled months, partial patterns

### Number Format System
- Right-click metric in sidebar → Excel-style format popup
- Format types: General, Number, Currency, Percentage, Custom
- Configurable: decimal places, commas, prefix/suffix
- Stored per metric key in `_metric_formats` and `_metric_format_details`
- Applied in sidebar display and PowerPoint insertion
- Floating point cleanup: values within 0.0001 of integer → converted to int

## Known Issues / Technical Debt
1. ~~**ingestion_engine.py is ~3000 lines**~~ — ✅ Resolved in v16.0 (split into `config/`, `ui/`, `mapper/`, `engine/`)
2. ~~**Dead code**~~ — ✅ Resolved in v16.0 (~380 lines removed)
3. ~~**COM error handling**~~ — ✅ Resolved in v16.0 (structured logging, no bare excepts)
4. **Template auto-fill** — the one-click auto-fill from saved templates needs more testing with real data
5. **Chart data update** — implemented but not extensively tested with real PowerPoint charts
6. **Table data update** — implemented but not tested with real PowerPoint tables
7. **Undo** — uses `CommandBars.ExecuteMso("Undo")` which undoes globally, not per-shape
8. **Multiple files same platform** — works but platform config matching relies on exact sheet names

## Dependencies
```
beautifulsoup4    # HTML parsing
openpyxl          # Excel reading/writing
tkcalendar        # Calendar date picker
Pillow            # Image processing (thumbnail preview)
python-pptx       # PowerPoint file manipulation
pandas            # Data pipeline
pywin32           # Windows COM automation (optional, Windows only)
```

## File Inventory

### Entry Point
| File | Lines | Purpose |
|------|-------|---------|
| main.py | ~30 | Entry point: creates IngestionEngine, calls run() |

### UI Layer (`ui/`)
| File | Lines | Purpose |
|------|-------|---------|
| main_window.py | ~355 | Main window — file import, date range, Quick Run, export flow |
| review_view.py | ~270 | KPI cards, campaign detail, report generation (mixin) |
| client_wizard.py | ~270 | Campaign-to-client assignment wizard |
| settings_window.py | ~245 | Settings tabs (Platforms, Appearance, Dashboard Metrics, Export) |
| platform_setup.py | ~210 | Import sample file, assign column roles |
| utils.py | ~45 | enable_mousewheel() — universal scroll binding |

### Template Mapper (`mapper/`)
| File | Lines | Purpose |
|------|-------|---------|
| mapper_window.py | ~290 | PPTXWizard — 3-column mapper with embedded preview |
| slide_view.py | ~340 | Slide rendering, shape assignment, preview thumbnail |
| sidebar.py | ~235 | Metric sidebar with search, right-click format |
| query_builder.py | ~305 | Advanced query popup with visual pivot table |
| format_popup.py | ~135 | Excel-style number-format popup |

### Engine (`engine/`)
| File | Lines | Purpose |
|------|-------|---------|
| data_pipeline.py | ~175 | apply_platform_config(), filter_data_by_campaigns() |
| excel_writer.py | ~430 | Pandas-powered 3-sheet Excel output |
| excel_utils.py | ~190 | Excel formatting helpers |
| pptx_mapper.py | ~195 | PowerPoint scan, map, fill (python-pptx) |
| pptx_fill.py | ~200 | Shape-level PowerPoint fill logic |
| pptx_formats.py | ~185 | Number/date formatting for PowerPoint values |
| pptx_live.py | ~390 | COM live preview (chart, table, text, image) |
| query_resolver.py | ~215 | Pandas query engine for metric resolution |
| metrics_catalog.py | ~135 | get_available_metrics(), SPECIAL_METRICS |
| kpi.py | ~65 | KPI_METRICS, compute_kpis() |
| errors.py | ~20 | IngestionError, ParserError |

### Parsers (`parsers/`)
| File | Lines | Purpose |
|------|-------|---------|
| dictionary.py | ~170 | Metric alias matching, column classification |
| csv_parser.py | ~145 | CSV file ingestion |
| excel_parser.py | ~225 | Excel file ingestion with smart header detection |
| html_parser.py | ~275 | HTML table ingestion |

### Configuration (`config/`)
| File | Lines | Purpose |
|------|-------|---------|
| settings.py | ~55 | load_settings(), save_settings(), platform config |
| themes.py | ~25 | THEMES dict, get_theme() |
| paths.py | ~40 | All path constants (OUTPUT_DIR, TEMPLATES_DIR, etc.) |
| logging_setup.py | ~35 | Rotating file logging setup |

## Workflow (User-Facing)

```
1. Launch → Main Window
   ├── Add Files (CSV, XLSX, HTML)
   ├── Assign Platforms (dropdown per file)
   ├── Set Date Range (calendar pickers)
   └── Quick Run →

2. Client Assignment (maximized popup)
   ├── Enter client name
   ├── Check campaigns (grouped by platform)
   ├── Next → (loops for more clients)
   └── Finish →

3. Review (in main window, maximized)
   ├── KPI Cards (dynamic, pandas-powered)
   ├── Campaign Detail (expandable)
   ├── Generate Report → 
   │   ├── If templates exist → Template Selector (one-click auto-fill)
   │   └── If no templates → Template Mapper (full editor)
   ├── Next Client →
   └── Finish

4. Template Mapper (maximized, 3-column)
   ├── Left: Metrics sidebar (right-click to format)
   │   ├── Quick Fill (Client Name, Dates, Image, Custom Text)
   │   ├── Advanced Query Builder
   │   └── KPI Totals + Breakdown sections
   ├── Center: Slide Preview (PNG thumbnail, refreshes after changes)
   ├── Right: Shape list (assign, skip, multi-assign, clear)
   └── Bottom: Prev/Next, Save Mapping, Save & Fill
```

## Pending Features (Post-Refactor)
- [ ] Move template mapper to Settings (pre-build templates, AEs select from dropdown)
- [ ] .exe packaging with PyInstaller
- [ ] Cross-source data reconciliation
- [ ] Data quality validation (breakdown sums = campaign totals)
- [ ] Chart data filling with real PowerPoint charts (tested but needs validation)
- [ ] Table filling with real PowerPoint tables
- [ ] Campaign type detection/assignment
- [ ] Client name auto-detection from campaign names
- [ ] PowerPoint table auto-population from query builder
