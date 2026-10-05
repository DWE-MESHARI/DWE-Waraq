---
name: Browser-local document scope
description: Product boundary for presentation uploads, analysis, notes, and image extraction.
---

Keep uploaded presentations, extracted slide text, findings, notes, and note images in browser storage. Do not add server-side document processing or cloud persistence without a new user request. Text embedded in slide images is not automatically extracted and this limitation should remain clear.

**Why:** The user specified a local-first review workflow and the OCR limitation.

**How to apply:** Keep parsing and saved review data client-side, retain user-triggered HTML/JSON exports for sharing, and disclose the image-text limitation in the interface and reports.
