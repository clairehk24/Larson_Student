LARSON HTML PROOF OF CONCEPT
============================

START
Open index.html in a current desktop browser. The prototype does not require an internet connection or a web server.

WHAT THIS CONTAINS
- A persistent, collapsible simulation menu.
- A right-side viewer based on the uploaded navigation/viewer concept.
- The Rise POC hierarchy and the downloadable files from the Rise web export.
- Rise-inspired visual settings: orange accent, dark navy lesson header, rounded cards, light navigation, and compact lesson headers.
- Responsive navigation for smaller screens.
- Search, previous/next controls, open-page control, and browser-based completion tracking.

ADD A SIMULATION
1. Duplicate pages/_simulation-template.html and rename it.
2. Replace the bracketed placeholders with approved simulation content.
3. Copy related downloads into assets/.
4. Add a new item in data/navigation.js with a unique id, title, and relative path.

CONTENT WORKFLOW
The Rise web export stores course text in an encoded runtime-data.js file. This prototype uses clean standalone HTML instead of Rise's minified runtime. For a larger build, Word manuscripts can be converted into this repeatable page structure while the navigation remains unchanged.

VERIFY MANUSCRIPT TEXT
Run `python scripts/validate_manuscript_text.py` after editing manuscript-based pages. The check compares the reader-facing HTML blocks with their source Word paragraphs and fails if wording, spelling, punctuation, capitalization, or ordering changes.

NOTES
- The two lessons labeled Placeholder retain their copied Rise content and display a review notice.
- No external libraries or packaged font files are used.
- Completion is stored in the learner's browser only. LMS reporting would require a SCORM/xAPI wrapper or LMS-specific integration.
