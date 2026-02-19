You are generating code for a Chrome Extension (Manifest V3). 
The extension is internal-only and must run fully offline with NO external network calls and NO CDN usage.

Goal:
- When the user clicks the extension’s toolbar icon, the extension extracts the page’s main content and downloads it as a .docx file.
- Ignore content inside <header> and <footer> tags completely.
- Preserve formatting as much as reasonably possible: headings (h1–h6), paragraphs, bold/italic/underline, links, unordered/ordered lists, and tables.
- Do NOT include navigation chrome, ads, or unrelated page elements beyond the main content.

Content selection rules:
1) Prefer <main> if present.
2) Else prefer <article> if present.
3) Else select the largest visible text-containing container element (by text length) that is NOT inside header/footer.
4) Always exclude any nodes that are descendants of <header> or <footer>.
5) Always exclude any nodes that are visibility hidden or display none via css.

DOCX generation:
- Output must be a real .docx file.
- Use a locally bundled library to generate DOCX in the browser (example: “docx” JS library). Do not reference external URLs.
- Bundle dependencies into the extension (e.g., via a small build step). Include the built output in /dist and reference it from the extension files.
- The generated .docx should include a title at top (use document.title) and then the extracted content.

Extension behavior:
- Clicking the extension icon triggers the export (no popup UI required).
- Use a background service worker that injects/executes a content script on the active tab and requests the extracted structured content.
- The background script then builds the docx and triggers a download using the Chrome downloads API.
- Name the file using a safe slug of document.title plus a timestamp (YYYY-MM-DD_HH-mm).

Technical requirements:
- Manifest V3.
- Minimal permissions: activeTab, scripting, downloads.
- Vanilla JavaScript (no frameworks).
- Clean, minimal code with comments where needed.
- Handle SPAs reasonably: extraction happens at click-time from the current DOM.

Deliverables:
1) Provide the full folder/file structure.
2) Complete manifest.json.
3) background.js (service worker).
4) content.js (DOM extraction + conversion into a structured intermediate format).
5) A small doc builder module that maps the structured format into docx objects.
6) Build instructions (one-liner is best) for bundling dependencies locally (e.g. esbuild/rollup), but the final answer must include the final source files and the expected /dist outputs.

Implementation detail:
- The content script should traverse DOM and emit a JSON structure like:
  - blocks: [{type:"heading", level:1..6, runs:[{text,bold,italic,underline,href?}]}, {type:"paragraph", runs:[...]}, {type:"list", ordered:boolean, items:[blocks...]}, {type:"table", rows:[cells...]}, ...]
- Then background builds the docx from that JSON.
- Ensure header/footer descendants are skipped during traversal.

Output:
- Provide code only (no fluff), and ensure it is runnable with minimal setup.
