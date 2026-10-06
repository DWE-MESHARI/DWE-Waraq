---
name: Browser-local document scope
description: Product boundary for presentation uploads, analysis, notes, and image extraction.
---

Keep uploaded presentations, extracted slide text, findings, notes, note images, and embedded slide-image assets in browser storage. Do not add server-side document processing or cloud persistence without a new user request. Extract images only from the presentation package; never fetch externally linked images automatically. Text embedded in slide images is not automatically extracted and this limitation should remain clear.

**Why:** The user specified a local-first review workflow and selected images embedded in PPTX slides for display. Supporting those images must not change the privacy boundary or imply OCR support.

**How to apply:** Keep parsing and saved review data client-side, retain user-triggered HTML/JSON exports for sharing, preserve older saved analyses, and disclose the image-text limitation in the interface and reports.
