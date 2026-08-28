# Larson Student Version — WCAG 2.0 Level AA Audit

Audit date: August 28, 2026  
Target: WCAG 2.0, conformance level AA  
Scope: the course shell, Introduction, Simulations 1–32, and the 221 files linked for download

## Executive summary

The HTML course interface and all 218 Word activities have been remediated for the defects found in this audit. All three PDFs now have unique titles and document language, and the SCAT6 Instructions PDF is tagged. Two licensed, prefilled SCAT6 case PDFs remain untagged and therefore still prevent a whole-course WCAG 2.0 Level AA conformance claim until they are tagged and manually reviewed in Acrobat.

This is a code-assisted audit, not a certification. W3C notes that conformance evaluation requires both automated/semi-automated checks and manual review. The manual tests in this report must be completed after remediation.

## Scope and method

The review covered:

- `index.html`, `pages/introduction.html`, and `pages/simulation-1.html` through `pages/simulation-32.html`
- Shared behavior and presentation in `app.js`, `styles.css`, and `pages/page.css`
- All 218 DOCX and 3 PDF files under `assets/downloads`
- Static checks for document language, titles, headings, image alternatives, form names, iframe titles, duplicate IDs, local link targets, and malformed text
- Computed color contrast using the WCAG relative-luminance formula
- Source-level keyboard and focus-order review at the responsive breakpoint
- OOXML inspection of Word document metadata, heading styles, drawings, and table-header markup
- Preliminary PDF tag/metadata inspection

The reusable page template was inspected but is not counted as a live course page.

## Remediation results

### 1. Orange primary-control contrast — remediated

**Success criterion:** 1.4.3 Contrast (Minimum), Level AA  
**Severity:** High  
**Original evidence:** White (`#ffffff`) against the original orange accent (`#d65e00`) had a contrast ratio of **3.85:1**. WCAG 2.0 requires at least 4.5:1 for normal text.

Affected uses include:

- The small white text in `.complete-button`
- The small white simulation number in the active `.nav-number`

**Change made:** The accent is now `#c45100`, which provides **4.64:1** against white. Focus outlines also use the solid passing accent instead of a translucent color.

### 2. Closed mobile navigation keyboard focus — remediated

**Success criteria:** 2.4.3 Focus Order and 2.4.7 Focus Visible, Level A/AA  
**Severity:** High  
**Original evidence:** At widths of 980px or less, `.sidebar` was moved off screen while its controls remained in the tab order.

**Changes made:**

- The closed responsive sidebar now uses `inert` and `aria-hidden="true"`.
- Opening it removes those states and moves focus to search.
- Escape, scrim activation, menu toggling, and lesson selection restore focus appropriately.
- Viewport changes synchronize the sidebar’s interaction state.

### 3. Navigation completion state — remediated

**Success criterion:** 1.4.1 Use of Color, Level A  
**Severity:** Medium  
**Original evidence:** `.nav-link.is-complete .nav-status` changed the same check icon from gray to green.

**Change made:** Incomplete items display an empty circle; completed items display a circle with a checkmark. The programmatic label continues to change between “Not complete” and “Complete.”

### 4. Word document title metadata — remediated

**Success criterion:** 2.4.2 Page Titled, Level A  
**Severity:** High due to scale  
**Original evidence:** **209 of 218 DOCX files** had an empty Dublin Core title in `docProps/core.xml`.

**Change made:** All **218 of 218 DOCX files** now have a descriptive title derived from the document’s activity heading.

### 5. Word image alternatives — remediated

**Success criterion:** 1.1.1 Non-text Content, Level A  
**Severity:** High  
**Original evidence:** All **6 drawings across 5 DOCX files** lacked alternative-text data:

- `simulation-1/presimulation-activity-2-anatomy-labeling.docx` — 1 image
- `simulation-7/presimulation-activity-4-anatomy.docx` — 2 images
- `simulation-8/presimulation-activity-3-anatomy-labeling.docx` — 1 image
- `simulation-15/postsimulation-activity-1-pulse-and-heart-rate.docx` — 1 image
- `simulation-16/presimulation-activity-4-power-wheel.docx` — 1 image

**Change made:** All six drawings now have task-appropriate descriptions. Labeling exercises identify the diagram and its labeling purpose without invalidating the exercise; informative figures describe the relevant labels and relationships.

### 6. Word heading hierarchy — remediated

**Success criterion:** 1.3.1 Info and Relationships, Level A  
**Severity:** Medium  
**Original evidence:** **19 files began with Heading 2** and had no preceding Heading 1.

**Change made:** All **218 of 218 DOCX files** now begin with Heading 1.

### 7. Word table structure — remediated; manual verification retained

**Success criteria:** 1.3.1 Info and Relationships and 1.3.2 Meaningful Sequence, Level A  
**Severity:** High pending manual classification  
**Original evidence:** There were **44 tables in 25 DOCX files**, and none contained a marked header row or accessible table metadata.

**Changes made:** Every table now has a descriptive caption, a table description, and a marked header row. Word Accessibility Checker and screen-reader review are still required because several worksheet layouts use tables for form-like entry areas.

- For data tables, identify a header row, enable “Repeat as header row,” avoid split/merged cells where possible, and verify screen-reader navigation.
- Replace layout tables with ordinary paragraphs, tabs, or styles when reading order cannot be made reliable.

### 8. PDF accessibility — partially remediated; blocking

**Success criteria:** 1.3.1 Info and Relationships and 1.3.2 Meaningful Sequence, Level A  
**Severity:** High  
**Evidence:** Full parsing confirmed no `/StructTreeRoot` in:

- `simulation-21/hcp-scat6-rain-shoemaker.pdf`
- `simulation-22/96-hour-scat6-rain-shoemaker.pdf`

All three PDFs now have unique title metadata and `en-US` document language. `simulation-21/scat6-instructions.pdf` retains its structure tree. The two case PDFs remain untagged.

Required remediation: run Acrobat Pro’s **Autotag document** on the two case PDFs, then manually correct headings, tables, form fields, alternate text, and reading order. Because these are complex licensed clinical forms, automatic tagging alone is not sufficient and a homegrown reformatted transcript is not an appropriate substitute.

## Advisory issues

These issues are not, by themselves, confirmed WCAG 2.0 failures but should be cleaned up:

- Decorative file badges and download arrows are now consistently hidden from assistive technology on all simulation pages.
- Dynamic lesson loads now update a polite live status, and the iframe/browser titles continue to update.
- The skip-link target can now receive programmatic focus and has a visible focus outline.
- The “Open page” button opens a new browsing context. A visible or accessible warning would make the behavior more predictable even though WCAG 2.0 does not require it at Level AA in every circumstance.

## Checks that passed

The static HTML review found:

- A declared `lang="en"` on every live HTML document
- A unique document title on every live HTML document
- Exactly one `h1` and no skipped heading levels on every live HTML document
- Alternative text on the course logo
- An accessible name for the search field and every button
- A descriptive title on the lesson iframe
- Native keyboard-capable elements for links, buttons, and navigation disclosures
- A visible skip link
- Visible focus styling in the shared CSS
- No duplicate IDs in static markup
- All **221 local download links resolve to existing files**
- No videos, audio, or time-based media requiring captions or audio description in the HTML pages
- Reduced-motion styling for users who request it

## Required manual verification

Complete these tests after the confirmed failures are fixed:

1. Keyboard-only test in Edge and Chrome at desktop width and at 980px or below, including forward/reverse tab order, menu open/close, search, all disclosures, lesson changes, downloads, completion, and previous/next controls.
2. NVDA + Edge test for landmarks, headings, lesson-navigation state, iframe entry/exit, completion state, search results, and lesson changes.
3. Text resize to 200% and browser zoom to 200%, checking for clipped text, obscured controls, and loss of content or function.
4. Windows High Contrast/forced-colors check, especially active lesson and completion indicators.
5. Word’s Accessibility Checker on every DOCX after the batch metadata/structure corrections.
6. Acrobat Pro full check plus manual tag-tree and reading-order review for all three PDFs.
7. Manual content review of link purpose, plain-language instructions, abbreviation expansion, and whether image alternatives communicate the intended learning task.

## Remediation order

1. Autotag and manually remediate the two untagged SCAT6 case PDFs in Acrobat Pro.
2. Run the manual browser and NVDA verification listed above.
3. Run Word Accessibility Checker on the remediated activities, paying special attention to form-like tables.
4. Run Acrobat’s full check and manual tag-tree/reading-order review on all three PDFs.

## References

- [Web Content Accessibility Guidelines (WCAG) 2.0](https://www.w3.org/TR/WCAG20/)
- [W3C WCAG-EM conformance evaluation overview](https://www.w3.org/WAI/test-evaluate/conformance/wcag-em/)
- [W3C accessibility evaluation report template](https://www.w3.org/WAI/test-evaluate/report-template/)
- [W3C evaluation tools overview](https://www.w3.org/WAI/test-evaluate/tools/)
