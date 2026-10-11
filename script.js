/*
 * وَرَق — مراجعة عروض PowerPoint محليًا في المتصفح (نسخة ثابتة)
 * لا يحتاج بناءً ولا خادمًا: افتح index.html مباشرة.
 *
 * الأقسام:
 *   1) محلل النصوص (قواعد الرصد والنفي)   2) قارئ ملفات PPTX
 *   3) التخزين المحلي IndexedDB            4) التصدير (JSON / تقرير HTML)
 *   5) الواجهة والتفاعل
 */
(function () {
'use strict';


/* ================= محلل النصوص (قواعد الرصد) ================= */
const CATEGORY_RULES = [
    ["السلامة", /خطر|سلامة|حاجز|ساتر|كبار السن|الأطفال|الاطفال|إصابة|اصابة|سطح/],
    ["الكهرباء والتكييف", /كهرب|أفياش|افياش|مكيف|تكييف|تمديد كهربائي|صندوق الكهرباء/],
    ["المياه والسباكة", /مياه|الماء|خزان|مواسير|سباكة|تسرب|مغسلة|صرف صحي|حمام/],
    ["الإنشاءات", /خرسانة|خرساني|صبة|صبّة|جدار|طوب|شروخ|شقوق|تشققات|أسوار|اسوار|مخطط|إنشائي|انشائي/],
    ["التشطيبات", /بلاط|دهان|ديكور|زجاج|درابزين|أبواب|ابواب|أسقف|اسقف|واجهات|شبابيك/],
    ["التواصل مع العملاء", /عميل|عملاء|بروشور|التنويه|المعلومات|شكوى|شكاوى|شكاوي/],
    ["التشغيل والمعاينة", /زيارة|زيارات|معاينة|فحص|مقاول|استلام|تنظيف|مخلفات|مستودع/],
];
const HIGH_WORDS = "خطر|إصابة|اصابة|تشققات|شروخ|تسرب|مكشوف|مكسور|متصدع|انهيار|غير آمن|غير امن|صعق";
const ISSUE_WORDS = "مشكلة|مشاكل|شكوى|شكاوى|شكاوي|غير مكتمل|نقص|مفقود|غير مناسب|غير مطابقة|غير مطابق|اختلاف|تفاوت|تأخير|مخلفات|تشققات|شروخ|تسرب|خطر|مكسور|مائلة|مائل|ضعيفة|ضعيف|قصور|خلل|عيب|عيوب";
// عبارات "غياب" هي بحد ذاتها مشكلة (عدم وجود حاجز)، لكنها تُلغى إن كان الغائب هو المشكلة نفسها.
const ABSENCE_ISSUE = "عدم وجود|لا يوجد|لا توجد|لايمكن|لا يمكن";
const HIGH_RISK_TERMS = new RegExp(HIGH_WORDS);
const ISSUE_TERMS = new RegExp(`${ISSUE_WORDS}|${ABSENCE_ISSUE}`);
/**
 * نفي الخطر: «لا يوجد خطر»، «دون تسرب»، «خالٍ من الشروخ»، «لم تظهر أي مشاكل».
 * تُحذف هذه العبارات قبل اختبار كلمات الخطر حتى لا تُصنَّف خطأً كمشكلة.
 */
const NEGATED_RISK = new RegExp(`(?<![\\u0621-\\u064A])(?:و?(?:لا|لم|ليس|ليست)(?:\\s+(?:يوجد|توجد|يظهر|تظهر|يلاحظ|تلاحظ|يكن|تكن|نلاحظ|نجد|يُلاحظ|تُلاحظ))?|عدم\\s+(?:وجود|ظهور|ملاحظة)|دون(?:\\s+وجود)?|بدون|خال(?:ية|يه|ٍ|ي)?\\s+من)\\s+(?:أي\\s+|اي\\s+|أية\\s+)?(?:ال)?(?:${ISSUE_WORDS})`, "g");
function neutralizeNegation(text) {
    return text.replace(NEGATED_RISK, " ");
}
const LABEL_PATTERN = /(الملاحظة|السبب المحتمل|الإجراء المقترح|الاجراء المقترح)\s*[:：-]?\s*/g;
const LABEL_HEADER = /^(الملاحظة|السبب المحتمل|الإجراء المقترح|الاجراء المقترح)$/;
/** القسم الأكثر تطابقًا مع النص (عند التعادل: الأسبق في القواعد). */
function categoryFor(text) {
    let best = "ملاحظات عامة", bestScore = 0;
    for (const [name, pattern] of CATEGORY_RULES) {
        const score = (text.match(new RegExp(pattern.source, "g")) || []).length;
        if (score > bestScore) { best = name; bestScore = score; }
    }
    return best;
}
function severityFor(text, kind) {
    if (kind === "توصية")
        return "منخفض";
    const effective = neutralizeNegation(text);
    if (HIGH_RISK_TERMS.test(effective))
        return "مرتفع";
    if (ISSUE_TERMS.test(effective))
        return "متوسط";
    return "منخفض";
}
function cleanText(value) {
    return value.replace(/\s+/g, " ").replace(/^[\s:：،-]+|[\s،]+$/g, "").trim();
}
/** بصمة نصية قصيرة وثابتة (djb2) لبناء معرّفات لا تتصادم عند تشابه بدايات النصوص. */
function hashText(value) {
    let hash = 5381;
    for (let index = 0; index < value.length; index += 1) {
        hash = ((hash << 5) + hash + value.charCodeAt(index)) | 0;
    }
    return (hash >>> 0).toString(36);
}

/* ---------- الحالات: ملاحظة + سبب + إجراء تُجمَّع في حالة واحدة ---------- */
const ROLE_KIND = { observation: "ملاحظة", cause: "ملاحظة", action: "توصية", sentence: "مشكلة" };
const SEVERITY_RANK = { "مرتفع": 0, "متوسط": 1, "منخفض": 2 };
const STATUS_LABELS = { open: "مفتوحة", progress: "قيد المتابعة", done: "مغلقة" };

function roleForLabel(label) {
    if (label.includes("السبب")) return "cause";
    if (label.includes("الإجراء") || label.includes("الاجراء")) return "action";
    return "observation";
}

/** يرصد مؤشرات شريحة واحدة ويربطها بمعرّف حالة (caseId) ودور (role). */
function analyzeSlideFindings(slide) {
    const out = [];
    let caseNo = 0;
    const newCaseId = (seed) => `${slide.number}-c${++caseNo}-${hashText(seed)}`;
    const push = (caseId, role, text) => {
        const cleaned = cleanText(text);
        if (cleaned.length < 8) return;
        const kind = ROLE_KIND[role];
        out.push({
            id: `${caseId}-${role}-${out.length}`, caseId, role,
            slideNumber: slide.number, section: "", kind,
            severity: severityFor(cleaned, kind), text: cleaned, evidence: cleaned,
        });
    };

    // 1) الجداول ذات رأس «الملاحظة | السبب المحتمل | الإجراء المقترح»: كل صف حالة.
    let prose = slide.text;
    for (const table of slide.tables ?? []) {
        const header = (table[0] ?? []).map(cleanText);
        const columns = header
            .map((cell, column) => ({ column, label: cell.match(LABEL_HEADER)?.[1] }))
            .filter((item) => item.label);
        if (!columns.length) continue;
        prose = prose.replace(tableToText(table), "\n");
        for (const row of table.slice(1)) {
            const first = cleanText(row[columns[0].column] ?? "");
            const caseId = newCaseId(first || row.join("|"));
            for (const { column, label } of columns) push(caseId, roleForLabel(label), row[column] ?? "");
        }
    }

    // 2) النص ذو التسميات: «الملاحظة» تفتح حالة جديدة، والسبب والإجراء يلتحقان بها.
    const matches = Array.from(prose.matchAll(LABEL_PATTERN));
    let current = null;
    matches.forEach((match, index) => {
        const start = (match.index ?? 0) + match[0].length;
        const end = matches[index + 1]?.index ?? prose.length;
        const body = cleanText(prose.slice(start, end));
        if (body.length < 8) return;
        const role = roleForLabel(match[1]);
        if (role === "observation" || !current || current.roles.has(role)) {
            current = { id: newCaseId(body), roles: new Set() };
        }
        current.roles.add(role);
        push(current.id, role, body);
    });

    // 3) جمل حرة فيها كلمات مشكلة (لم يسبق التقاطها بتسمياتها)
    const known = out.map((item) => item.text);
    prose.replace(/([.!؟؛])\s+/g, "$1\n").split(/\n+|\s\|\s/).map(cleanText)
        .filter((part) => part.length >= 20 && ISSUE_TERMS.test(neutralizeNegation(part)))
        .filter((part) => !known.some((text) => part.includes(text)))
        .slice(0, 6)
        .forEach((part) => push(newCaseId(part), "sentence", part));

    // قسم موحّد لكل حالة: من نص الملاحظة، وإلا من نص الحالة كاملًا
    const byCase = new Map();
    out.forEach((item) => { (byCase.get(item.caseId) ?? byCase.set(item.caseId, []).get(item.caseId)).push(item); });
    for (const items of byCase.values()) {
        const observation = items.find((item) => item.role === "observation" || item.role === "sentence") ?? items[0];
        let section = categoryFor(observation.text);
        if (section === "ملاحظات عامة") section = categoryFor(items.map((item) => item.text).join(" "));
        items.forEach((item) => { item.section = section; });
    }
    return out;
}

/** يبني قائمة الحالات من المؤشرات. المؤشرات القديمة (بلا caseId) تصبح كل منها حالة. */
function buildCases(findings, statuses = {}) {
    const map = new Map();
    for (const item of findings) {
        const key = item.caseId || item.id;
        if (!map.has(key)) map.set(key, { id: key, slideNumber: item.slideNumber, section: item.section, findings: [] });
        map.get(key).findings.push(item);
    }
    return Array.from(map.values()).map((c) => {
        const role = (r) => c.findings.find((f) => f.role === r);
        const legacy = c.findings.length === 1 && !c.findings[0].role;
        const observation = role("observation") ?? role("sentence") ?? (legacy && c.findings[0].kind !== "توصية" ? c.findings[0] : undefined);
        const cause = role("cause");
        const action = role("action") ?? (legacy && c.findings[0].kind === "توصية" ? c.findings[0] : undefined);
        const scored = [observation, cause].filter(Boolean);
        const severity = (scored.length ? scored : c.findings)
            .map((f) => f.severity).sort((a, b) => SEVERITY_RANK[a] - SEVERITY_RANK[b])[0];
        const first = c.findings[0];
        return {
            ...c, observation, cause, action, severity,
            manual: Boolean(first.manual), blockId: first.blockId, owner: first.owner || '', due: first.due || '',
            title: (observation ?? cause ?? action ?? c.findings[0]).text,
            status: statuses[c.id] || "open",
        };
    });
}

/** ملخص تنفيذي للعرض كله. */
function summarize(analysis) {
    const cases = buildCases(analysis.findings, analysis.statuses || {});
    const bySeverity = { "مرتفع": 0, "متوسط": 0, "منخفض": 0 };
    const byStatus = { open: 0, progress: 0, done: 0 };
    const bySection = new Map();
    const bySlide = new Map();
    for (const c of cases) {
        bySeverity[c.severity] += 1;
        byStatus[c.status] += 1;
        bySection.set(c.section, (bySection.get(c.section) ?? 0) + 1);
        bySlide.set(c.slideNumber, (bySlide.get(c.slideNumber) ?? 0) + 1);
    }
    const openHigh = cases.filter((c) => c.severity === "مرتفع" && c.status !== "done").length;
    const openMedium = cases.filter((c) => c.severity === "متوسط" && c.status !== "done").length;
    const risk = openHigh >= 3 ? "مرتفع" : (openHigh >= 1 || openMedium >= 5) ? "متوسط" : "منخفض";
    return {
        cases, bySeverity, byStatus, risk, openHigh,
        sections: Array.from(bySection.entries()).sort((a, b) => b[1] - a[1]),
        slides: Array.from(bySlide.entries()).sort((a, b) => b[1] - a[1] || a[0] - b[0]),
        priorities: cases.filter((c) => c.status !== "done")
            .sort((a, b) => SEVERITY_RANK[a.severity] - SEVERITY_RANK[b.severity] || a.slideNumber - b.slideNumber).slice(0, 5),
        progress: cases.length ? Math.round(byStatus.done / cases.length * 100) : 0,
    };
}

function tableToText(table) {
    return table.map((row) => row.join(" | ")).join("\n");
}

function analyzeSlides(fileName, slides, imageAssets = {}, imageStats = undefined) {
    const records = slides.map((slide) => ({
        number: slide.number,
        title: slide.title || `الشريحة ${slide.number}`,
        text: slide.text,
        notes: slide.notes || undefined,
        tables: slide.tables?.length ? slide.tables : undefined,
        imageIds: slide.imageIds ?? [],
        imagesSkipped: slide.imagesSkipped || 0,
    }));
    return {
        id: crypto.randomUUID(), fileName, createdAt: new Date().toISOString(),
        slideCount: slides.length, slides: records, imageAssets, imageStats,
        findings: slides.flatMap(analyzeSlideFindings), statuses: {}, notes: [],
    };
}

/* ================= قارئ ملفات PowerPoint (للاستيراد) ================= */
const POWERPOINT_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const RELATIONSHIP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SVG_IMAGE_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const PACKAGE_RELATIONSHIP_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const MAX_FILE_BYTES = 150 * 1024 * 1024;
const MAX_SLIDE_XML_BYTES = 25 * 1024 * 1024;
const MAX_TOTAL_IMAGE_BYTES = 150 * 1024 * 1024;
function getZipEntries(bytes) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const lowerBound = Math.max(0, bytes.byteLength - 65_557);
    let endRecordOffset = -1;
    for (let offset = bytes.byteLength - 22; offset >= lowerBound; offset -= 1) {
        if (view.getUint32(offset, true) === 0x06054b50) {
            endRecordOffset = offset;
            break;
        }
    }
    if (endRecordOffset < 0) {
        throw new Error("لم يتم العثور على بنية ملف PowerPoint صالحة.");
    }
    const entryCount = view.getUint16(endRecordOffset + 10, true);
    const centralDirectoryOffset = view.getUint32(endRecordOffset + 16, true);
    if (entryCount === 0xffff ||
        centralDirectoryOffset === 0xffffffff) {
        throw new Error("هذا الملف يستخدم تنسيق ضغط غير مدعوم أو يحتوي على عناصر كثيرة جدًا.");
    }
    const decoder = new TextDecoder("utf-8");
    const entries = [];
    let offset = centralDirectoryOffset;
    for (let index = 0; index < entryCount; index += 1) {
        if (offset + 46 > bytes.byteLength ||
            view.getUint32(offset, true) !== 0x02014b50) {
            throw new Error("تعذر قراءة محتويات ملف PowerPoint.");
        }
        const fileNameLength = view.getUint16(offset + 28, true);
        const extraLength = view.getUint16(offset + 30, true);
        const commentLength = view.getUint16(offset + 32, true);
        const nameStart = offset + 46;
        const nameEnd = nameStart + fileNameLength;
        const name = decoder.decode(bytes.subarray(nameStart, nameEnd));
        entries.push({
            name,
            compressionMethod: view.getUint16(offset + 10, true),
            compressedSize: view.getUint32(offset + 20, true),
            uncompressedSize: view.getUint32(offset + 24, true),
            localHeaderOffset: view.getUint32(offset + 42, true),
        });
        offset = nameEnd + extraLength + commentLength;
    }
    return entries;
}
async function readZipEntry(bytes, entry, maxUncompressedBytes = MAX_SLIDE_XML_BYTES) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const localOffset = entry.localHeaderOffset;
    if (localOffset + 30 > bytes.byteLength ||
        view.getUint32(localOffset, true) !== 0x04034b50) {
        throw new Error(`تعذر فتح الملف الداخلي: ${entry.name}`);
    }
    if (entry.uncompressedSize > maxUncompressedBytes) {
        throw new Error(`حجم العنصر الداخلي أكبر من الحد المسموح: ${entry.name}`);
    }
    const nameLength = view.getUint16(localOffset + 26, true);
    const extraLength = view.getUint16(localOffset + 28, true);
    const dataStart = localOffset + 30 + nameLength + extraLength;
    const dataEnd = dataStart + entry.compressedSize;
    if (dataEnd > bytes.byteLength) {
        throw new Error(`بيانات الشريحة غير مكتملة: ${entry.name}`);
    }
    const compressed = bytes.subarray(dataStart, dataEnd);
    if (entry.compressionMethod === 0)
        return compressed.slice();
    if (entry.compressionMethod !== 8) {
        throw new Error("طريقة ضغط هذا الملف غير مدعومة.");
    }
    if (typeof DecompressionStream === "undefined") {
        throw new Error("المتصفح لا يدعم فك ضغط ملفات PowerPoint الحديثة.");
    }
    try {
        const input = new Blob([compressed.slice().buffer]).stream();
        const decompressed = input.pipeThrough(new DecompressionStream("deflate-raw"));
        return new Uint8Array(await new Response(decompressed).arrayBuffer());
    }
    catch {
        throw new Error(`تعذر فك ضغط بيانات الشريحة: ${entry.name}`);
    }
}
function parseXml(xml, label) {
    const document = new DOMParser().parseFromString(xml, "application/xml");
    if (document.getElementsByTagName("parsererror").length > 0) {
        throw new Error(`تعذر تحليل ملف ${label}.`);
    }
    return document;
}
function getOrderedSlidePaths(fileNames, presentationXml, relationshipsXml) {
    const fallback = Array.from(fileNames).sort((a, b) => {
        const aNumber = Number(a.match(/slide(\d+)\.xml$/)?.[1] ?? 0);
        const bNumber = Number(b.match(/slide(\d+)\.xml$/)?.[1] ?? 0);
        return aNumber - bNumber;
    });
    if (!presentationXml || !relationshipsXml)
        return fallback;
    try {
        const presentation = parseXml(presentationXml, "ترتيب الشرائح");
        const relationships = parseXml(relationshipsXml, "روابط الشرائح");
        const targets = new Map();
        for (const relationship of Array.from(relationships.getElementsByTagNameNS(PACKAGE_RELATIONSHIP_NS, "Relationship"))) {
            const id = relationship.getAttribute("Id");
            const target = relationship.getAttribute("Target");
            if (!id || !target)
                continue;
            const normalizedTarget = target.startsWith("/")
                ? target.slice(1)
                : target.startsWith("ppt/")
                    ? target
                    : `ppt/${target}`;
            targets.set(id, normalizedTarget);
        }
        const ordered = Array.from(presentation.getElementsByTagNameNS(POWERPOINT_NS, "sldId"))
            .map((slide) => {
            const relationId = slide.getAttributeNS(RELATIONSHIP_NS, "id");
            return relationId ? targets.get(relationId) : undefined;
        })
            .filter((path) => Boolean(path && fileNames.has(path)));
        return ordered.length > 0 ? ordered : fallback;
    }
    catch {
        return fallback;
    }
}
function paragraphsOf(container) {
    return Array.from(container.getElementsByTagNameNS(DRAWING_NS, "p"))
        .map((paragraph) => Array.from(paragraph.getElementsByTagNameNS(DRAWING_NS, "t"))
        .map((node) => node.textContent?.replace(/\s+/g, " ").trim() ?? "")
        .join("")
        .trim())
        .filter(Boolean);
}
function directChildren(parent, namespace, name) {
    return Array.from(parent.children).filter((child) => child.localName === name && child.namespaceURI === namespace);
}
function readTable(table) {
    return directChildren(table, DRAWING_NS, "tr").map((row) => directChildren(row, DRAWING_NS, "tc").map((cell) => paragraphsOf(cell).join(" ")));
}
function isTitleShape(shape) {
    const placeholder = shape.getElementsByTagNameNS(POWERPOINT_NS, "ph")[0];
    const type = placeholder?.getAttribute("type");
    return type === "title" || type === "ctrTitle";
}
/**
 * يمشي على الشريحة بترتيب ظهور العناصر: الأشكال النصية (بما فيها داخل المجموعات)
 * والجداول (p:graphicFrame ← a:tbl). عنوان الشريحة يؤخذ من حاوية العنوان إن وُجدت.
 */
function extractText(xml, slideNumber) {
    const document = parseXml(xml, `الشريحة ${slideNumber}`);
    const blocks = [];
    const tables = [];
    let titleText;
    const visit = (node) => {
        if (node.namespaceURI === POWERPOINT_NS && node.localName === "sp") {
            const text = paragraphsOf(node).join("\n");
            if (text) {
                blocks.push(text);
                if (!titleText && isTitleShape(node))
                    titleText = text.split("\n")[0];
            }
            return;
        }
        if (node.namespaceURI === DRAWING_NS && node.localName === "tbl") {
            const grid = readTable(node);
            if (grid.some((row) => row.some(Boolean))) {
                tables.push(grid);
                blocks.push(grid.map((row) => row.join(" | ")).join("\n"));
            }
            return;
        }
        for (const child of Array.from(node.children))
            visit(child);
    };
    visit(document.documentElement);
    const text = blocks.length > 0
        ? blocks.join("\n")
        : Array.from(document.getElementsByTagNameNS(DRAWING_NS, "t"))
            .map((node) => node.textContent?.replace(/\s+/g, " ").trim() ?? "")
            .filter(Boolean)
            .join("\n");
    const title = (titleText ?? blocks[0]?.split("\n")[0] ?? "").slice(0, 120)
        || `الشريحة ${slideNumber}`;
    return { number: slideNumber, title, text, tables: tables.length ? tables : undefined };
}
/** نص ملاحظات المتحدث (ppt/notesSlides). نتجاهل حاوية رقم الشريحة وصورتها. */
function extractNotes(xml, slideNumber) {
    const document = parseXml(xml, `ملاحظات الشريحة ${slideNumber}`);
    return Array.from(document.getElementsByTagNameNS(POWERPOINT_NS, "sp"))
        .filter((shape) => {
        const type = shape.getElementsByTagNameNS(POWERPOINT_NS, "ph")[0]?.getAttribute("type");
        return type === "body" || type === undefined || type === null;
    })
        .map((shape) => paragraphsOf(shape).join("\n"))
        .filter(Boolean)
        .join("\n");
}
function relationshipPathFor(sourcePath) {
    const separator = sourcePath.lastIndexOf("/");
    const directory = separator >= 0 ? sourcePath.slice(0, separator) : "";
    const fileName = separator >= 0 ? sourcePath.slice(separator + 1) : sourcePath;
    return `${directory}/_rels/${fileName}.rels`;
}
function resolvePackagePath(sourcePath, target) {
    const isRootRelative = target.startsWith("/") || target.startsWith("ppt/");
    const parts = isRootRelative
        ? []
        : sourcePath.split("/").slice(0, -1);
    const targetParts = (target.startsWith("/") ? target.slice(1) : target)
        .replace(/\\/g, "/")
        .split("/");
    for (const part of targetParts) {
        if (!part || part === ".")
            continue;
        if (part === "..") {
            parts.pop();
            continue;
        }
        parts.push(part);
    }
    return parts.join("/");
}
function parseRelationships(xml) {
    const relationships = new Map();
    if (!xml)
        return relationships;
    const document = parseXml(xml, "روابط الصور");
    for (const relationship of Array.from(document.getElementsByTagNameNS(PACKAGE_RELATIONSHIP_NS, "Relationship"))) {
        const id = relationship.getAttribute("Id");
        const target = relationship.getAttribute("Target");
        const type = relationship.getAttribute("Type");
        if (!id || !target || !type)
            continue;
        relationships.set(id, {
            target,
            type,
            targetMode: relationship.getAttribute("TargetMode") || "",
        });
    }
    return relationships;
}
function parseContentTypes(xml) {
    const defaults = new Map();
    const overrides = new Map();
    if (!xml)
        return { defaults, overrides };
    const document = parseXml(xml, "أنواع ملفات العرض");
    for (const node of Array.from(document.getElementsByTagName("*"))) {
        if (node.localName === "Default") {
            const extension = node.getAttribute("Extension")?.toLowerCase();
            const contentType = node.getAttribute("ContentType");
            if (extension && contentType)
                defaults.set(extension, contentType);
        }
        else if (node.localName === "Override") {
            const partName = node.getAttribute("PartName");
            const contentType = node.getAttribute("ContentType");
            if (partName && contentType)
                overrides.set(partName.replace(/^\//, ""), contentType);
        }
    }
    return { defaults, overrides };
}
function fallbackImageType(path) {
    const extension = path.split(".").pop()?.toLowerCase();
    const types = {
        avif: "image/avif",
        bmp: "image/bmp",
        gif: "image/gif",
        jpeg: "image/jpeg",
        jpg: "image/jpeg",
        png: "image/png",
        svg: "image/svg+xml",
        tif: "image/tiff",
        tiff: "image/tiff",
        webp: "image/webp",
        wmf: "image/x-wmf",
        emf: "image/x-emf",
    };
    return (extension && types[extension]) || "application/octet-stream";
}
function getImageType(path, contentTypes) {
    const declaredType = contentTypes.overrides.get(path)
        || contentTypes.defaults.get(path.split(".").pop()?.toLowerCase() || "");
    return declaredType && /^image\/[a-z0-9.+-]+$/i.test(declaredType)
        ? declaredType
        : fallbackImageType(path);
}
function fileNameFromPath(path) {
    return path.split(/[\\/]/).pop() || path;
}

/* ---------- الصور الأساسية فقط ----------
 * الافتراضي: صور الشرائح الفعلية (p:pic) دون الشعارات والخلفيات والزخارف والأيقونات
 * والصور المكررة على أغلب الشرائح. خيار «كل الصور» يعيد السلوك الشامل. */
const IMAGE_NAME_HINTS = /logo|شعار|icon|أيقونة|ايقونة|watermark|علامة مائية|background|خلفية|bullet|decor|زخرف|divider|separator|footer|header/i;
const NON_DISPLAYABLE = /\.(emf|wmf)$/i;
const DEFAULT_SLIDE_SIZE = { cx: 9144000, cy: 6858000 };
const SKIP_REASONS = {
    decor: "شعارات وزخارف", small: "صور صغيرة/أيقونات", strip: "خطوط وفواصل", edge: "رأس/تذييل الشريحة",
    background: "خلفيات", repeated: "متكررة على أغلب الشرائح", unsupported: "صيغة غير قابلة للعرض (EMF/WMF)", missing: "مفقودة داخل الملف",
};

function readSlideSize(presentationXml) {
    if (!presentationXml) return DEFAULT_SLIDE_SIZE;
    try {
        const node = parseXml(presentationXml, "حجم الشريحة").getElementsByTagNameNS(POWERPOINT_NS, "sldSz")[0];
        const cx = Number(node?.getAttribute("cx")), cy = Number(node?.getAttribute("cy"));
        return cx > 0 && cy > 0 ? { cx, cy } : DEFAULT_SLIDE_SIZE;
    } catch { return DEFAULT_SLIDE_SIZE; }
}

function blipRelationshipId(blip, relationships) {
    const svgBlip = blip.getElementsByTagNameNS(SVG_IMAGE_NS, "svgBlip")[0];
    const svgId = svgBlip?.getAttributeNS(RELATIONSHIP_NS, "embed") || svgBlip?.getAttribute("r:embed");
    const regularId = blip.getAttributeNS(RELATIONSHIP_NS, "embed") || blip.getAttributeNS(RELATIONSHIP_NS, "link")
        || blip.getAttribute("r:embed") || blip.getAttribute("r:link");
    return svgId && relationships.has(svgId) ? svgId : regularId;
}

function insideGroup(node) {
    for (let parent = node.parentElement; parent; parent = parent.parentElement) {
        if (parent.localName === "grpSp") return true;
    }
    return false;
}

/** يجمع صور الشريحة بترتيب ظهورها مع موضع كل صورة وحجمها إن وُجد. */
function collectPictures(xml, slideNumber, relationships, allImages) {
    const document = parseXml(xml, `صور الشريحة ${slideNumber}`);
    const list = [];
    if (allImages) {
        for (const blip of Array.from(document.getElementsByTagNameNS(DRAWING_NS, "blip"))) {
            const rid = blipRelationshipId(blip, relationships);
            if (rid) list.push({ rid, name: "", box: null });
        }
        return list;
    }
    for (const pic of Array.from(document.getElementsByTagNameNS(POWERPOINT_NS, "pic"))) {
        const blip = pic.getElementsByTagNameNS(DRAWING_NS, "blip")[0];
        const rid = blip && blipRelationshipId(blip, relationships);
        if (!rid) continue;
        const props = pic.getElementsByTagNameNS(POWERPOINT_NS, "cNvPr")[0];
        const name = `${props?.getAttribute("name") ?? ""} ${props?.getAttribute("descr") ?? ""} ${props?.getAttribute("title") ?? ""}`;
        let box = null;
        const spPr = directChildren(pic, POWERPOINT_NS, "spPr")[0];
        const xfrm = spPr && directChildren(spPr, DRAWING_NS, "xfrm")[0];
        const ext = xfrm && directChildren(xfrm, DRAWING_NS, "ext")[0];
        const off = xfrm && directChildren(xfrm, DRAWING_NS, "off")[0];
        if (ext && !insideGroup(pic)) {
            box = { x: Number(off?.getAttribute("x") ?? 0), y: Number(off?.getAttribute("y") ?? 0), cx: Number(ext.getAttribute("cx")), cy: Number(ext.getAttribute("cy")) };
            if (!(box.cx > 0 && box.cy > 0)) box = null;
        }
        list.push({ rid, name, box });
    }
    return list;
}

/** سبب استبعاد الصورة (أو null إن كانت صورة أساسية). */
function skipReasonFor(picture, target, size, textLength) {
    if (IMAGE_NAME_HINTS.test(picture.name)) return "decor";
    if (NON_DISPLAYABLE.test(target || "")) return "unsupported";
    const box = picture.box;
    if (!box) return null; // موضعها موروث من القالب: نعدّها صورة محتوى
    const area = (box.cx * box.cy) / (size.cx * size.cy);
    const ratio = box.cx / box.cy;
    if (area >= 0.85 && textLength >= 40) return "background";
    if (area < 0.025) return "small";
    if (ratio > 8 || ratio < 0.125) return "strip";
    const height = box.cy / size.cy, top = box.y / size.cy, bottom = (box.y + box.cy) / size.cy;
    if (height < 0.15 && area < 0.08 && (top < 0.12 || bottom > 0.88)) return "edge";
    return null;
}

async function loadImageAssets(plan, byName, bytes, contentTypes) {
    const imageAssets = {};
    let total = 0;
    for (const item of plan) {
        if (imageAssets[item.assetId]) continue;
        const fileName = fileNameFromPath(item.target || item.assetId);
        if (item.kind === "missing") {
            imageAssets[item.assetId] = { id: item.assetId, fileName: `صورة الشريحة ${item.slideNumber}`, contentType: "application/octet-stream", error: "لم يتم العثور على علاقة الصورة داخل ملف العرض." };
        } else if (item.kind === "external") {
            imageAssets[item.assetId] = { id: item.assetId, fileName, contentType: "image/*", externalUrl: item.target };
        } else {
            const entry = byName.get(item.assetId);
            if (!entry) {
                imageAssets[item.assetId] = { id: item.assetId, fileName, contentType: fallbackImageType(item.assetId), error: "لم يتم العثور على ملف الصورة المضمّن داخل العرض." };
                continue;
            }
            if (total + entry.uncompressedSize > MAX_TOTAL_IMAGE_BYTES) {
                throw new Error("يتجاوز الحجم الإجمالي للصور المضمّنة 150 ميغابايت؛ جرّب الوضع الافتراضي (الصور الأساسية فقط).");
            }
            const contentType = getImageType(item.assetId, contentTypes);
            const imageBytes = await readZipEntry(bytes, entry, MAX_TOTAL_IMAGE_BYTES);
            total += imageBytes.byteLength;
            imageAssets[item.assetId] = { id: item.assetId, fileName, contentType, blob: new Blob([imageBytes.slice().buffer], { type: contentType }) };
        }
    }
    return imageAssets;
}

async function readPowerPoint(file, options = {}) {
    const allImages = Boolean(options.allImages);
    if (!file.name.toLowerCase().endsWith(".pptx")) throw new Error("يرجى اختيار ملف PowerPoint بامتداد ‎.pptx.");
    if (file.size > MAX_FILE_BYTES) throw new Error("حجم الملف أكبر من 150 ميغابايت.");

    const bytes = new Uint8Array(await file.arrayBuffer());
    const entries = getZipEntries(bytes);
    const byName = new Map(entries.map((entry) => [entry.name, entry]));
    const slideNames = new Set(entries.filter((e) => /^ppt\/slides\/slide\d+\.xml$/.test(e.name)).map((e) => e.name));
    if (slideNames.size === 0) throw new Error("لم يتم العثور على شرائح قابلة للقراءة في هذا الملف.");

    const readTextEntry = async (name) => {
        const entry = byName.get(name);
        return entry ? new TextDecoder("utf-8").decode(await readZipEntry(bytes, entry)) : undefined;
    };
    const [presentationXml, presentationRels] = await Promise.all([
        readTextEntry("ppt/presentation.xml"), readTextEntry("ppt/_rels/presentation.xml.rels"),
    ]);
    const contentTypes = parseContentTypes(await readTextEntry("[Content_Types].xml"));
    const size = readSlideSize(presentationXml);
    const orderedPaths = getOrderedSlidePaths(slideNames, presentationXml, presentationRels);

    // المرحلة 1: النص والجداول والملاحظات، وجمع كل صورة مرشّحة دون قراءة بياناتها
    const slides = [];
    const candidates = []; // {slideIndex, slideNumber, assetId, kind, target, reason}
    for (let index = 0; index < orderedPaths.length; index += 1) {
        const path = orderedPaths[index];
        const entry = byName.get(path);
        if (!entry) continue;
        const xml = new TextDecoder("utf-8").decode(await readZipEntry(bytes, entry));
        const slideNumber = index + 1;
        const relationships = parseRelationships(await readTextEntry(relationshipPathFor(path)));
        const extracted = extractText(xml, slideNumber);

        let notes = "";
        const notesRelation = Array.from(relationships.values()).find((item) => item.type.toLowerCase().endsWith("/notesslide"));
        if (notesRelation) {
            try {
                const notesXml = await readTextEntry(resolvePackagePath(path, notesRelation.target));
                if (notesXml) notes = extractNotes(notesXml, slideNumber);
            } catch { notes = ""; }
        }

        const slideIndex = slides.length;
        for (const picture of collectPictures(xml, slideNumber, relationships, allImages)) {
            const relationship = relationships.get(picture.rid);
            if (!relationship) {
                candidates.push({ slideIndex, slideNumber, assetId: `missing:${path}:${picture.rid}`, kind: "missing", reason: allImages ? null : "missing" });
                continue;
            }
            if (!relationship.type.toLowerCase().endsWith("/image")) continue;
            const external = relationship.targetMode.toLowerCase() === "external";
            const assetId = external ? `external:${relationship.target}` : resolvePackagePath(path, relationship.target);
            candidates.push({
                slideIndex, slideNumber, assetId, kind: external ? "external" : "embedded", target: relationship.target,
                reason: allImages ? null : skipReasonFor(picture, relationship.target, size, extracted.text.length),
            });
        }
        slides.push({ ...extracted, notes: notes || undefined, imageIds: [], imagesSkipped: 0 });
    }

    // المرحلة 2: الصورة التي تظهر على أغلب الشرائح (شعار/إطار) تُستبعد
    if (!allImages) {
        const usage = new Map();
        for (const c of candidates) {
            if (!usage.has(c.assetId)) usage.set(c.assetId, new Set());
            usage.get(c.assetId).add(c.slideNumber);
        }
        for (const c of candidates) {
            const used = usage.get(c.assetId).size;
            if (!c.reason && used >= 3 && used > 0.3 * slides.length) c.reason = "repeated";
        }
    }

    // المرحلة 3: اعتماد الصور الأساسية وقراءة بياناتها فقط
    const kept = [];
    const skipped = {};
    for (const c of candidates) {
        const slide = slides[c.slideIndex];
        if (c.reason) {
            skipped[c.reason] = (skipped[c.reason] ?? 0) + 1;
            slide.imagesSkipped += 1;
        } else if (!slide.imageIds.includes(c.assetId)) {
            slide.imageIds.push(c.assetId);
            kept.push(c);
        }
    }
    const imageAssets = await loadImageAssets(kept, byName, bytes, contentTypes);
    const imageStats = {
        mode: allImages ? "all" : "main",
        found: candidates.length,
        kept: kept.length,
        skipped,
        skippedTotal: Object.values(skipped).reduce((sum, n) => sum + n, 0),
    };
    return { slides, imageAssets, imageStats };
}

/* ================= التخزين المحلي (IndexedDB) ================= */
const DB_NAME = 'waraq-analysis';
const DB_VERSION = 3;
const STORE = 'presentations';
/** الصور مخزّنة هنا كـ Blob مستقل، بمفتاح [معرّف التحليل، معرّف الأصل]، حتى لا يُعاد كتابتها مع كل ملاحظة. */
const ASSET_STORE = 'assets';
/** المشاريع (النموذج الأساسي): كل مشروع شرائح وعناصر. معرّف المشروع هو نفسه معرّف الصور في مخزن الأصول. */
const PROJECT_STORE = 'projects';
function openDatabase() {
    if (typeof indexedDB === 'undefined') {
        return Promise.reject(new Error('التخزين المحلي غير متاح في هذا المتصفح.'));
    }
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(DB_NAME, DB_VERSION);
        request.onupgradeneeded = () => {
            const db = request.result;
            if (!db.objectStoreNames.contains(STORE)) {
                db.createObjectStore(STORE, { keyPath: 'id' });
            }
            if (!db.objectStoreNames.contains(PROJECT_STORE)) {
                db.createObjectStore(PROJECT_STORE, { keyPath: 'id' });
            }
            if (!db.objectStoreNames.contains(ASSET_STORE)) {
                db.createObjectStore(ASSET_STORE, { keyPath: ['analysisId', 'assetId'] });
            }
        };
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}
function runTransaction(stores, mode, work, failure) {
    return openDatabase().then((db) => new Promise((resolve, reject) => {
        const tx = db.transaction(stores, mode);
        let result;
        const request = work(tx);
        if (request)
            request.onsuccess = () => { result = request.result; };
        tx.oncomplete = () => { db.close(); resolve(result); };
        tx.onerror = () => { db.close(); reject(tx.error || new Error(failure)); };
        tx.onabort = () => { db.close(); reject(tx.error || new Error(failure)); };
    }));
}
async function getAnalyses() {
    const result = await runTransaction([STORE], 'readonly', tx => tx.objectStore(STORE).getAll(), 'تم إيقاف قراءة التحليلات.');
    return result.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}
/** يفصل الـ Blob عن بيانات التحليل: يبقى في السجل وصف الأصل فقط (stored: true). */
function splitAssets(analysis) {
    const blobs = [];
    const imageAssets = analysis.imageAssets && {};
    for (const [assetId, asset] of Object.entries(analysis.imageAssets ?? {})) {
        const { blob, ...meta } = asset;
        if (blob) {
            blobs.push({ analysisId: analysis.id, assetId, blob });
            imageAssets[assetId] = { ...meta, stored: true };
        }
        else {
            imageAssets[assetId] = meta;
        }
    }
    return { record: { ...analysis, imageAssets }, blobs };
}
/** حفظ تحليل جديد بكل صوره. للتعديلات اللاحقة (الملاحظات) استخدم saveAnalysisMeta. */
async function saveAnalysis(analysis) {
    const { record, blobs } = splitAssets(analysis);
    await runTransaction([STORE, ASSET_STORE], 'readwrite', tx => {
        tx.objectStore(STORE).put(record);
        const assets = tx.objectStore(ASSET_STORE);
        for (const item of blobs)
            assets.put(item);
    }, 'تعذر حفظ التغييرات.');
    return record;
}
/** حفظ سجل التحليل (النصوص والملاحظات) دون إعادة كتابة الصور. */
async function saveAnalysisMeta(analysis) {
    const { record } = splitAssets(analysis);
    await runTransaction([STORE], 'readwrite', tx => { tx.objectStore(STORE).put(record); }, 'تعذر حفظ التغييرات.');
}
async function removeAnalysis(id) {
    await runTransaction([STORE, ASSET_STORE], 'readwrite', tx => {
        tx.objectStore(STORE).delete(id);
        tx.objectStore(ASSET_STORE).delete(IDBKeyRange.bound([id, ''], [id, '￿']));
    }, 'تعذر حذف التحليل.');
}
async function getAssetBlobs(analysisId) {
    const rows = await runTransaction([ASSET_STORE], 'readonly', tx => tx.objectStore(ASSET_STORE).getAll(IDBKeyRange.bound([analysisId, ''], [analysisId, '￿'])), 'تعذر قراءة الصور.');
    return Object.fromEntries(rows.map(row => [row.assetId, row.blob]));
}



/* ---------- المشاريع ---------- */
async function getProjects() {
    const rows = await runTransaction([PROJECT_STORE], 'readonly', tx => tx.objectStore(PROJECT_STORE).getAll(), 'تعذر قراءة المشاريع.');
    return rows.sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
}
async function saveProject(project, quiet = false) {
    const record = { ...project };
    delete record._rev;
    await runTransaction([PROJECT_STORE], 'readwrite', tx => { tx.objectStore(PROJECT_STORE).put(record); }, 'تعذر حفظ المشروع.');
    if (!quiet && typeof queueSync === 'function') queueSync();
}
async function putAssetBlob(projectId, assetId, blob) {
    await runTransaction([ASSET_STORE], 'readwrite', tx => { tx.objectStore(ASSET_STORE).put({ analysisId: projectId, assetId, blob }); }, 'تعذر حفظ الصورة.');
}
async function deleteProject(id) {
    await runTransaction([PROJECT_STORE, ASSET_STORE], 'readwrite', tx => {
        tx.objectStore(PROJECT_STORE).delete(id);
        tx.objectStore(ASSET_STORE).delete(IDBKeyRange.bound([id, ''], [id, '\uffff']));
    }, 'تعذر حذف المشروع.');
}
/** حذف سجل التحليل القديم فقط (الصور باقية لأن معرّفها صار معرّف المشروع). */
async function removeLegacyRecord(id) {
    await runTransaction([STORE], 'readwrite', tx => { tx.objectStore(STORE).delete(id); }, 'تعذر تحديث السجل القديم.');
}
async function copyAssetBlobs(fromId, toId) {
    const blobs = await getAssetBlobs(fromId);
    for (const [assetId, blob] of Object.entries(blobs)) await putAssetBlob(toId, assetId, blob);
}

async function getAssetBlob(projectId, assetId) {
    const row = await runTransaction([ASSET_STORE], 'readonly', tx => tx.objectStore(ASSET_STORE).get([projectId, assetId]), 'تعذر قراءة الصورة.');
    return row?.blob;
}
async function deleteAssetBlob(projectId, assetId) {
    await runTransaction([ASSET_STORE], 'readwrite', tx => { tx.objectStore(ASSET_STORE).delete([projectId, assetId]); }, 'تعذر حذف الصورة.');
}

/* ================= نموذج المشروع (الشرائح والعناصر) ================= */
const uid = () => (crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(16).slice(2));
const CATEGORY_NAMES = [...CATEGORY_RULES.map(([name]) => name), 'ملاحظات عامة'];
const SEVERITY_CHOICES = ['مرتفع', 'متوسط', 'منخفض'];
const BLOCK_LABELS = { text: 'نص', finding: 'مؤشر', image: 'صورة', comment: 'تعليق', table: 'جدول' };

function newBlock(type, extra = {}) {
    const base = { id: uid(), type };
    if (type === 'text') return { ...base, style: 'p', text: '', ...extra };
    if (type === 'finding') return { ...base, text: '', cause: '', action: '', severity: 'auto', category: 'auto', status: 'open', owner: '', due: '', ...extra };
    if (type === 'image') return { ...base, assetId: '', caption: '', ...extra };
    if (type === 'comment') return { ...base, text: '', author: '', createdAt: new Date().toISOString(), ...extra };
    if (type === 'table') return { ...base, rows: [['', ''], ['', '']], ...extra };
    return base;
}
function newSlide(title = '', blocks = []) { return { id: uid(), title, notes: '', blocks }; }

const TEMPLATES = {
    blank: () => [newSlide('الشريحة 1')],
    review: () => [
        newSlide('غلاف المشروع', [newBlock('text', { style: 'h', text: 'تقرير المراجعة الميدانية' }), newBlock('text', { text: 'اكتب نطاق المراجعة وتاريخها والجهات المشاركة.' })]),
        newSlide('ملخص الوضع', [newBlock('text', { style: 'ul', text: 'نسبة الإنجاز العامة\nأهم ما رُصد أثناء الزيارة' })]),
        newSlide('الملاحظات الميدانية', [newBlock('finding', { text: '', cause: '', action: '' })]),
        newSlide('التوصيات والخطوات التالية', [newBlock('text', { style: 'ul', text: '' })]),
    ],
};

function newProject({ name, owner = '', location = '', description = '', template = 'blank' }) {
    const now = new Date().toISOString();
    return {
        id: uid(), name: name.trim(), owner: owner.trim(), location: location.trim(), description: description.trim(),
        createdAt: now, updatedAt: now, slides: (TEMPLATES[template] || TEMPLATES.blank)(),
        statuses: {}, imageAssets: {}, history: [], version: 1,
    };
}

/** نص الشريحة الذي يحلله النظام: عناصر النص والجداول (بلا العنوان ولا المؤشرات اليدوية). */
function slideBodyText(slide) {
    const parts = [];
    for (const b of slide.blocks) {
        if (b.type === 'text' && b.text.trim()) parts.push(b.text.trim());
        else if (b.type === 'table') parts.push(tableToText(b.rows));
    }
    return parts.join('\n');
}

/** يحوّل المشروع إلى «تحليل» بنفس البنية التي تعرضها واجهة التحليل. */
function projectToAnalysis(project) {
    const findings = [];
    const comments = [];
    const statuses = { ...(project.statuses || {}) };
    const imageAssets = {};
    const slides = project.slides.map((slide, index) => {
        const number = index + 1;
        const title = slide.title.trim() || `الشريحة ${number}`;
        const body = slideBodyText(slide);
        const tables = slide.blocks.filter(b => b.type === 'table').map(b => b.rows);
        findings.push(...analyzeSlideFindings({ number, text: body, tables }));
        const imageIds = [];
        for (const b of slide.blocks) {
            if (b.type === 'image' && b.assetId) {
                imageIds.push(b.assetId);
                const meta = project.imageAssets?.[b.assetId] || { id: b.assetId, contentType: 'image/*' };
                imageAssets[b.assetId] = { ...meta, fileName: b.caption || meta.fileName || 'صورة', stored: true };
            } else if (b.type === 'finding') {
                const parts = [['observation', b.text], ['cause', b.cause], ['action', b.action]].filter(([, t]) => cleanText(t || '').length >= 2);
                if (!parts.length) continue;
                const caseId = `m-${b.id}`;
                const all = parts.map(([, t]) => t).join(' ');
                const section = b.category && b.category !== 'auto' ? b.category : categoryFor(all);
                parts.forEach(([role, t]) => {
                    const cleaned = cleanText(t);
                    findings.push({
                        id: `${caseId}-${role}`, caseId, role, slideNumber: number, section, kind: ROLE_KIND[role],
                        severity: SEVERITY_CHOICES.includes(b.severity) ? b.severity : severityFor(cleaned, ROLE_KIND[role]),
                        text: cleaned, evidence: cleaned, manual: true, blockId: b.id, owner: b.owner || '', due: b.due || '',
                    });
                });
                statuses[caseId] = b.status || 'open';
            } else if (b.type === 'comment' && b.text.trim()) {
                comments.push({ id: b.id, text: b.text.trim(), slideNumber: number, author: b.author || '', createdAt: b.createdAt });
            }
        }
        const text = [title, body].filter(Boolean).join('\n');
        return { number, title, text, notes: slide.notes?.trim() || undefined, tables: tables.length ? tables : undefined, imageIds, id: slide.id };
    });
    return {
        id: project.id, fileName: project.name, owner: project.owner, location: project.location, description: project.description,
        createdAt: project.createdAt, updatedAt: project.updatedAt, slideCount: slides.length, slides, imageAssets,
        imageStats: project.imageStats, findings, statuses, notes: comments, history: project.history || [],
    };
}

/** تحليل مخزَّن مؤقتًا لكل مشروع حتى يتغير (state.rev). */
const analysisCache = new Map();
function analysisOf(project) {
    const key = project._rev || 0;
    const hit = analysisCache.get(project.id);
    if (hit && hit.key === key && hit.updatedAt === project.updatedAt) return hit;
    const analysis = projectToAnalysis(project);
    const summary = summarize(analysis);
    const entry = { key, updatedAt: project.updatedAt, analysis, summary };
    analysisCache.set(project.id, entry);
    return entry;
}

/* ---------- الاستيراد من PowerPoint أو من تحليل قديم ---------- */
function linesToBlocks(text, title) {
    const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
    if (lines[0] && title && lines[0] === title.trim()) lines.shift();
    return lines.map(line => newBlock('text', { text: line }));
}
function stripTables(text, tables) {
    return (tables || []).reduce((t, table) => t.replace(tableToText(table), '\n'), text);
}
/** يبني مشروعًا من الشرائح المستخرجة. الصور تُحفظ في مخزن الأصول بمعرّف المشروع. */
function slidesToProject({ name, slides, imageAssets = {}, imageStats, notes = [], statuses = {}, id, createdAt, owner = '' }) {
    const now = new Date().toISOString();
    const project = {
        id: id || uid(), name, owner, location: '', description: '', createdAt: createdAt || now, updatedAt: now,
        slides: [], statuses, imageAssets: {}, history: [], version: 1, imageStats,
    };
    for (const s of slides) {
        const blocks = linesToBlocks(stripTables(s.text || '', s.tables), s.title);
        for (const table of s.tables || []) blocks.push(newBlock('table', { rows: table.map(r => r.map(c => String(c ?? ''))) }));
        for (const imageId of s.imageIds || []) {
            const asset = imageAssets[imageId];
            if (!asset || asset.error || asset.externalUrl) continue;
            const { blob, dataUrl, ...meta } = asset;
            project.imageAssets[imageId] = { ...meta, id: imageId, stored: true };
            blocks.push(newBlock('image', { assetId: imageId, caption: '' }));
        }
        project.slides.push({ id: uid(), title: s.title || `الشريحة ${s.number}`, notes: s.notes || '', blocks });
    }
    for (const n of notes) {
        const slide = project.slides[(n.slideNumber || 1) - 1] || project.slides[0];
        if (!slide) continue;
        slide.blocks.push(newBlock('comment', { text: n.text + (n.quote ? `\n«${n.quote}»` : ''), createdAt: n.createdAt }));
    }
    if (!project.slides.length) project.slides.push(newSlide('الشريحة 1'));
    return project;
}

async function dataUrlToBlob(dataUrl) { return (await fetch(dataUrl)).blob(); }

/** ترحيل تحليل قديم (مخزن presentations) إلى مشروع بنفس المعرّف، فتبقى صوره في مكانها. */
async function analysisToProject(analysis) {
    const assets = analysis.imageAssets || {};
    for (const [assetId, asset] of Object.entries(assets)) {
        if (asset.dataUrl && !asset.stored) {
            try { await putAssetBlob(analysis.id, assetId, await dataUrlToBlob(asset.dataUrl)); asset.stored = true; } catch { /* تُهمل الصورة التالفة */ }
        }
    }
    const project = slidesToProject({
        id: analysis.id, name: analysis.fileName.replace(/\.pptx$/i, ''), slides: analysis.slides.map(s => ({ ...s })),
        imageAssets: Object.fromEntries(Object.entries(assets).filter(([, a]) => a.stored)),
        imageStats: analysis.imageStats, notes: [], statuses: analysis.statuses || {}, createdAt: analysis.createdAt,
    });
    // الصور المرفقة بملاحظات قديمة تصبح عناصر صور
    for (const n of analysis.notes || []) {
        const slide = project.slides[(n.slideNumber || 1) - 1] || project.slides[0];
        slide.blocks.push(newBlock('comment', { text: n.text + (n.quote ? `\n«${n.quote}»` : ''), createdAt: n.createdAt }));
        if (n.imageDataUrl) {
            try {
                const assetId = uid();
                await putAssetBlob(project.id, assetId, await dataUrlToBlob(n.imageDataUrl));
                project.imageAssets[assetId] = { id: assetId, fileName: 'مرفق ملاحظة', contentType: 'image/*', stored: true };
                slide.blocks.push(newBlock('image', { assetId, caption: 'مرفق ملاحظة' }));
            } catch { /* يُتجاهل المرفق التالف */ }
        }
    }
    return project;
}

/* ---------- نسخة احتياطية قابلة للنقل بين الأجهزة ---------- */
async function projectToBackup(project) {
    const blobs = await getAssetBlobs(project.id).catch(() => ({}));
    const images = {};
    for (const [id, blob] of Object.entries(blobs)) images[id] = await blobToDataUrl(blob);
    const copy = { ...project }; delete copy._rev;
    return { format: 'waraq-project', version: 1, exportedAt: new Date().toISOString(), project: copy, images };
}
async function backupToProject(data) {
    if (!data || data.format !== 'waraq-project' || !data.project || !Array.isArray(data.project.slides)) throw new Error('الملف ليس نسخة احتياطية صالحة من وَرَق.');
    const project = { ...data.project, id: uid(), name: `${data.project.name} (مستورد)`, updatedAt: new Date().toISOString() };
    for (const [assetId, url] of Object.entries(data.images || {})) await putAssetBlob(project.id, assetId, await dataUrlToBlob(url));
    return project;
}

/* ================= أدوات العرض والتصدير ================= */
const severityTone = { مرتفع: 'severity-high', متوسط: 'severity-mid', منخفض: 'severity-low' };
function formatDate(value) {
    return new Intl.DateTimeFormat('ar', { day: 'numeric', month: 'long', year: 'numeric' }).format(new Date(value));
}
function blobToDataUrl(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result));
        reader.onerror = () => reject(reader.error);
        reader.readAsDataURL(blob);
    });
}
function downloadFile(name, content, type) {
    const url = URL.createObjectURL(new Blob([content], { type }));
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = name.replace(/[\\/:*?"<>|]+/g, '-');
    anchor.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function escapeHtml(text) {
    return String(text ?? '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char] || char);
}
const baseName = (analysis) => analysis.fileName.replace(/\.pptx$/i, '');

/** نسخة احتياطية كاملة (مع الصور) يمكن استيرادها في أي متصفح. */
async function downloadBackup(project) {
    const data = await projectToBackup(project);
    downloadFile(`${project.name}.waraq.json`, JSON.stringify(data), 'application/json;charset=utf-8');
}

/** CSV يفتح في Excel بالعربية (UTF-8 مع BOM): صف لكل حالة. */
function downloadCasesCsv(analysis) {
    const cases = summarize(analysis).cases
        .sort((a, b) => a.slideNumber - b.slideNumber || SEVERITY_RANK[a.severity] - SEVERITY_RANK[b.severity]);
    const cell = (value) => `"${String(value ?? '').replace(/"/g, '""')}"`;
    const rows = [['الشريحة', 'القسم', 'الأهمية', 'حالة المتابعة', 'الملاحظة', 'السبب المحتمل', 'الإجراء المقترح', 'المسؤول', 'الموعد', 'مصدر']]
        .concat(cases.map(c => [c.slideNumber, c.section, c.severity, STATUS_LABELS[c.status], c.observation?.text ?? '', c.cause?.text ?? '', c.action?.text ?? '', c.owner, c.due, c.manual ? 'يدوي' : 'تلقائي']));
    downloadFile(`${baseName(analysis)}-الحالات.csv`, '﻿' + rows.map(r => r.map(cell).join(',')).join('\r\n'), 'text/csv;charset=utf-8');
}

/** تقرير HTML مستقل ومرتب: ملخص، أولويات، حالات حسب القسم، الشرائح بصورها الأساسية، الملاحظات. */
async function downloadAnalysisReport(analysis) {
    const blobs = await getAssetBlobs(analysis.id).catch(() => ({}));
    const dataUrls = {};
    for (const [id, blob] of Object.entries(blobs)) dataUrls[id] = await blobToDataUrl(blob);
    const e = escapeHtml;
    const s = summarize(analysis);
    const sorted = (list) => list.slice().sort((a, b) => SEVERITY_RANK[a.severity] - SEVERITY_RANK[b.severity] || a.slideNumber - b.slideNumber);
    const tag = (severity) => `<span class="tag ${severityTone[severity]}">${e(severity)}</span>`;
    const caseHtml = (c) => `<article class="case ${severityTone[c.severity]}">
      <div class="meta">${tag(c.severity)}<span>شريحة ${c.slideNumber}</span><span>${e(c.section)}</span><span class="st st-${c.status}">${e(STATUS_LABELS[c.status])}</span>${c.owner ? `<span>المسؤول: ${e(c.owner)}</span>` : ''}${c.due ? `<span>الموعد: ${e(c.due)}</span>` : ''}</div>
      <h4>${e(c.observation?.text ?? c.title)}</h4>
      ${c.cause ? `<p><b>السبب المحتمل:</b> ${e(c.cause.text)}</p>` : ''}${c.action ? `<p class="action"><b>الإجراء المقترح:</b> ${e(c.action.text)}</p>` : ''}</article>`;
    const sections = s.sections.map(([name]) => {
        const list = sorted(s.cases.filter(c => c.section === name));
        return `<section><h3>${e(name)} <small>(${list.length})</small></h3>${list.map(caseHtml).join('')}</section>`;
    }).join('');
    const imagesFor = (slide) => (slide.imageIds ?? []).map(id => {
        const asset = analysis.imageAssets?.[id];
        const src = dataUrls[id] || asset?.dataUrl;
        return src && !asset?.error ? `<figure><img src="${e(src)}" alt="${e(asset?.fileName)}"><figcaption>${e(asset?.fileName)}</figcaption></figure>` : '';
    }).join('');
    const slidesHtml = analysis.slides.map(slide => {
        const count = s.slides.find(([n]) => n === slide.number)?.[1] ?? 0;
        return `<article class="slide"><div class="meta"><b>الشريحة ${slide.number}</b>${count ? `<span class="tag severity-mid">${count} حالة</span>` : ''}</div><h4>${e(slide.title)}</h4>
        <p>${e(slide.text).replace(/\n/g, '<br>')}</p>${slide.notes ? `<blockquote><b>ملاحظات المتحدث:</b><br>${e(slide.notes).replace(/\n/g, '<br>')}</blockquote>` : ''}<div class="imgs">${imagesFor(slide)}</div></article>`;
    }).join('');
    const notesHtml = analysis.notes.map(n => `<article class="slide"><div class="meta">${n.slideNumber ? `<b>الشريحة ${n.slideNumber}</b>` : '<b>عامة</b>'}<span>${e(formatDate(n.createdAt))}</span></div><p>${e(n.text)}</p>${n.quote ? `<blockquote>${e(n.quote)}</blockquote>` : ''}${n.imageDataUrl ? `<img class="note-img" src="${e(n.imageDataUrl)}" alt="مرفق">` : ''}</article>`).join('');
    const st = analysis.imageStats;
    const html = `<!doctype html><html lang="ar" dir="rtl"><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>تقرير ${e(analysis.fileName)}</title>
<style>:root{--o:#e96826;--ink:#303532;--muted:#69736e;--line:#f0e5db}*{box-sizing:border-box}body{font:16px/1.8 system-ui,'Segoe UI',Tahoma,sans-serif;max-width:980px;margin:32px auto;padding:0 20px;color:var(--ink)}
header{border-bottom:3px solid var(--o);padding-bottom:16px;margin-bottom:22px}h1{margin:0;font-size:28px}h2{margin:34px 0 10px;padding-bottom:6px;border-bottom:1px solid var(--line);font-size:21px}h3{margin:22px 0 8px;font-size:17px}h3 small{color:var(--muted);font-weight:400}h4{margin:6px 0;font-size:16px}
.muted{color:var(--muted);font-size:14px}.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}.card{border:1px solid var(--line);border-radius:12px;padding:12px;text-align:center}.card b{display:block;font-size:26px}
.bar{height:10px;border-radius:6px;background:#f1eee5;overflow:hidden;display:flex;margin:8px 0}.bar i{display:block;height:100%}.done{background:#5e8469}.progress{background:#d9a441}.openb{background:#d4c9b9}
.case,.slide{border:1px solid var(--line);border-radius:12px;padding:14px 16px;margin:10px 0;break-inside:avoid;border-right:5px solid #668c74}.case.severity-high{border-right-color:#ad5349}.case.severity-mid{border-right-color:#b77a37}
.meta{display:flex;gap:10px;flex-wrap:wrap;align-items:center;font-size:13px;color:var(--muted)}.tag{border-radius:20px;padding:1px 10px;font-size:12px;font-weight:600;background:#eaf0e7;color:#5e8469}.tag.severity-high{background:#f6e9e4;color:#ad5349}.tag.severity-mid{background:#f6efdf;color:#a56d2f}
p{margin:4px 0}.action{color:#2f6b52}blockquote{border-right:3px solid var(--o);background:#fff3ea;margin:10px 0;padding:8px 14px}.imgs{display:flex;flex-wrap:wrap;gap:10px}figure{margin:8px 0;flex:1 1 260px;max-width:420px}figure img,.note-img{width:100%;max-height:420px;object-fit:contain;border:1px solid var(--line);border-radius:8px}figcaption{font-size:12px;color:var(--muted)}
.st{padding:0 8px;border-radius:10px;background:#f1eee5}.st-done{background:#e0eee5;color:#2f6b52}.st-progress{background:#f6efdf;color:#a56d2f}
@media print{body{margin:0}h2{break-after:avoid}}@media(max-width:640px){.cards{grid-template-columns:repeat(2,1fr)}}</style>
<header><h1>${e(analysis.fileName)}</h1><p class="muted">${analysis.owner ? `المسؤول: ${e(analysis.owner)} · ` : ''}${analysis.location ? `الموقع: ${e(analysis.location)} · ` : ''}${analysis.slideCount} شريحة · أُنشئ ${e(formatDate(analysis.createdAt))}</p>${analysis.description ? `<p>${e(analysis.description)}</p>` : ''}
<p class="muted">المؤشرات التلقائية قواعد نصية أولية وتحتاج إلى مراجعتك؛ والمؤشرات اليدوية أدخلها المسؤول.</p></header>
<h2>الملخص التنفيذي</h2>
<div class="cards"><div class="card"><b>${s.cases.length}</b>حالة</div><div class="card"><b>${s.bySeverity['مرتفع']}</b>أهمية مرتفعة</div><div class="card"><b>${s.bySeverity['متوسط']}</b>أهمية متوسطة</div><div class="card"><b>${s.bySeverity['منخفض']}</b>أهمية منخفضة</div></div>
<p>مستوى الانتباه العام: <b>${e(s.risk)}</b> · أُغلق ${s.progress}% من الحالات (${s.byStatus.done} مغلقة، ${s.byStatus.progress} قيد المتابعة، ${s.byStatus.open} مفتوحة).</p>
<div class="bar"><i class="done" style="width:${s.cases.length ? s.byStatus.done / s.cases.length * 100 : 0}%"></i><i class="progress" style="width:${s.cases.length ? s.byStatus.progress / s.cases.length * 100 : 0}%"></i><i class="openb" style="flex:1"></i></div>
${st ? `<p class="muted">الصور: اعتُمد ${st.kept} من ${st.found}${st.skippedTotal ? ` (استُبعد ${st.skippedTotal}: ${Object.entries(st.skipped).map(([k, n]) => `${SKIP_REASONS[k] || k} ${n}`).join('، ')})` : ''}.</p>` : ''}
<h2>أولويات المراجعة</h2>${s.priorities.length ? s.priorities.map(caseHtml).join('') : '<p>لا توجد حالات مفتوحة.</p>'}
<h2>الحالات حسب القسم (${s.cases.length})</h2>${sections || '<p>لم تُرصد حالات.</p>'}
<h2>التعليقات (${analysis.notes.length})</h2>${notesHtml || '<p>لا توجد تعليقات.</p>'}
<h2>الشرائح (${analysis.slides.length})</h2>${slidesHtml}</html>`;
    downloadFile(`${baseName(analysis)}-تقرير.html`, html, 'text/html;charset=utf-8');
}

/* ================= تصدير المشروع إلى PowerPoint (.pptx) ================= */
/* ملف OOXML يُبنى يدويًا داخل المتصفح بلا مكتبات: ZIP بلا ضغط + شرائح RTL. */
const CRC_TABLE = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(bytes) { let c = 0xffffffff; for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
function zipStore(files) {
    const enc = new TextEncoder(), chunks = [], central = [];
    let offset = 0;
    const u16 = (v) => [v & 255, (v >> 8) & 255], u32 = (v) => [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255];
    for (const [name, data] of files) {
        const nameBytes = enc.encode(name), bytes = typeof data === 'string' ? enc.encode(data) : data, crc = crc32(bytes);
        const head = new Uint8Array([0x50, 0x4b, 3, 4, ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(bytes.length), ...u32(bytes.length), ...u16(nameBytes.length), ...u16(0)]);
        chunks.push(head, nameBytes, bytes);
        central.push(new Uint8Array([0x50, 0x4b, 1, 2, ...u16(20), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(bytes.length), ...u32(bytes.length), ...u16(nameBytes.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(offset)]), nameBytes);
        offset += head.length + nameBytes.length + bytes.length;
    }
    let size = 0; for (const c of central) size += c.length;
    const end = new Uint8Array([0x50, 0x4b, 5, 6, ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length), ...u32(size), ...u32(offset), ...u16(0)]);
    return new Blob([...chunks, ...central, end], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
}

const EMU = 914400, SLIDE_W = 13.333, SLIDE_H = 7.5, MARGIN = 0.6, CONTENT_W = SLIDE_W - MARGIN * 2, TOP = 1.55, BOTTOM = 7.15;
const emu = (inch) => Math.round(inch * EMU);
const xe = (s) => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const PX = { ink: '303532', title: '29453C', orange: 'E96826', muted: '6B6B62', high: 'AD5349', mid: 'B77A37', low: '5E8469', highBg: 'FBE9E5', midBg: 'FBF1DF', lowBg: 'E7F0E9', head: 'F3EFE4', line: 'D8CDBD' };

function pRun(text, { sz = 18, b = false, i = false, color = PX.ink } = {}) {
    return `<a:r><a:rPr lang="ar-SA" sz="${sz * 100}" b="${b ? 1 : 0}" i="${i ? 1 : 0}" dirty="0"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:rPr><a:t>${xe(text)}</a:t></a:r>`;
}
function para(runs, { bullet = false, algn = 'r', spaceAfter = 4 } = {}) {
    const ppr = bullet
        ? `<a:pPr marL="342900" indent="-342900" algn="${algn}" rtl="1"><a:spcAft><a:spcPts val="${spaceAfter * 100}"/></a:spcAft><a:buFont typeface="Arial"/><a:buChar char="&#8226;"/></a:pPr>`
        : `<a:pPr algn="${algn}" rtl="1"><a:spcAft><a:spcPts val="${spaceAfter * 100}"/></a:spcAft><a:buNone/></a:pPr>`;
    return `<a:p>${ppr}${runs}</a:p>`;
}
const linesFor = (text, perLine) => String(text || '').split('\n').reduce((n, l) => n + Math.max(1, Math.ceil(l.length / perLine)), 0);

function buildSlideItems(slide, project, sizes) {
    const items = [];
    for (const b of slide.blocks) {
        if (b.type === 'text' && b.text.trim()) {
            if (b.style === 'h') items.push({ h: linesFor(b.text, 62) * 0.46 + 0.12, sp: (id, y) => textBox(id, y, linesFor(b.text, 62) * 0.46 + 0.1, [para(pRun(b.text, { sz: 22, b: true, color: PX.orange }))]) });
            else {
                const lines = b.text.split('\n').filter(l => l.trim());
                const n = linesFor(lines.join('\n'), 88), hgt = n * 0.37 + lines.length * 0.05 + 0.12;
                items.push({ h: hgt, sp: (id, y) => textBox(id, y, hgt, lines.map(l => para(pRun(l, { sz: 18 }), { bullet: b.style === 'ul' }))) });
            }
        } else if (b.type === 'finding' && `${b.text}${b.cause}${b.action}`.trim()) {
            const sev = SEVERITY_CHOICES.includes(b.severity) ? b.severity : severityFor(cleanText(`${b.text} ${b.cause}`), 'مشكلة');
            const col = sev === 'مرتفع' ? PX.high : sev === 'متوسط' ? PX.mid : PX.low, bg = sev === 'مرتفع' ? PX.highBg : sev === 'متوسط' ? PX.midBg : PX.lowBg;
            const meta = [`الخطورة: ${sev}`, `الحالة: ${STATUS_LABELS[b.status] || ''}`, b.owner && `المسؤول: ${b.owner}`, b.due && `الموعد: ${b.due}`].filter(Boolean).join('  ·  ');
            const ps = [para(pRun(meta, { sz: 13, b: true, color: col }), { spaceAfter: 3 })];
            let n = 0.34;
            if (b.text.trim()) { ps.push(para(pRun(b.text, { sz: 18, b: true }))); n += linesFor(b.text, 84) * 0.37 + 0.05; }
            if (b.cause.trim()) { ps.push(para(pRun('السبب المحتمل: ', { sz: 16, b: true, color: PX.muted }) + pRun(b.cause, { sz: 16 }))); n += linesFor(b.cause, 78) * 0.33 + 0.05; }
            if (b.action.trim()) { ps.push(para(pRun('الإجراء المقترح: ', { sz: 16, b: true, color: PX.low }) + pRun(b.action, { sz: 16 }))); n += linesFor(b.action, 76) * 0.33 + 0.05; }
            const hgt = n + 0.22;
            items.push({ h: hgt + 0.1, sp: (id, y) => textBox(id, y, hgt, ps, { fill: bg, line: col, inset: 0.12, name: 'مؤشر' }) });
        } else if (b.type === 'image' && sizes[b.assetId]) {
            const { w, h: ih, rid } = sizes[b.assetId];
            const maxH = b.caption ? 4.3 : 4.8, k = Math.min(CONTENT_W / w, maxH / ih, 1 / 0.0001);
            const pw = w * k, ph = ih * k, total = ph + (b.caption ? 0.4 : 0) + 0.1;
            items.push({ h: total, rels: [b.assetId], sp: (id, y, ridOf) => picture(id, ridOf(b.assetId), (SLIDE_W - pw) / 2, y, pw, ph, b.caption) + (b.caption ? textBox(id + 1, y + ph + 0.02, 0.36, [para(pRun(b.caption, { sz: 13, color: PX.muted }), { algn: 'ctr' })]) : '') });
        } else if (b.type === 'table' && b.rows.length) {
            const cols = b.rows[0].length, perLine = Math.max(8, Math.floor(88 / cols));
            const rowH = b.rows.map(r => Math.max(1, ...r.map(c => linesFor(c, perLine))) * 0.27 + 0.16);
            const total = rowH.reduce((a, c) => a + c, 0);
            items.push({ h: total + 0.15, sp: (id, y) => table(id, y, b.rows, rowH, total) });
        } else if (b.type === 'comment' && b.text.trim()) {
            const t = `تعليق${b.author ? ` (${b.author})` : ''}: ${b.text}`, hgt = linesFor(t, 100) * 0.3 + 0.12;
            items.push({ h: hgt, sp: (id, y) => textBox(id, y, hgt, [para(pRun(t, { sz: 14, i: true, color: PX.muted }))]) });
        }
    }
    if (slide.notes?.trim()) {
        const t = `ملاحظات المتحدث: ${slide.notes.trim()}`, hgt = linesFor(t, 100) * 0.3 + 0.12;
        items.push({ h: hgt, sp: (id, y) => textBox(id, y, hgt, [para(pRun(t, { sz: 14, color: PX.muted }))]) });
    }
    return items;
}

function textBox(id, y, h, paras, { fill, line, inset = 0.05, name = 'نص', x = MARGIN, w = CONTENT_W } = {}) {
    const f = fill ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>` : '<a:noFill/>';
    const l = line ? `<a:ln w="19050"><a:solidFill><a:srgbClr val="${line}"/></a:solidFill></a:ln>` : '<a:ln><a:noFill/></a:ln>';
    const geom = fill ? 'roundRect' : 'rect';
    return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${xe(name)} ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="${geom}"><a:avLst/></a:prstGeom>${f}${l}</p:spPr><p:txBody><a:bodyPr wrap="square" lIns="${emu(inset)}" tIns="${emu(inset / 2 + 0.03)}" rIns="${emu(inset)}" bIns="${emu(inset / 2 + 0.03)}" rtlCol="1" anchor="t"><a:noAutofit/></a:bodyPr><a:lstStyle/>${paras.join('')}</p:txBody></p:sp>`;
}
function picture(id, rid, x, y, w, h, descr) {
    return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="صورة ${id}" descr="${xe(descr || '')}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
}
function table(id, y, rows, rowH, total, x = MARGIN, w = CONTENT_W, sz = 14) {
    const cols = rows[0].length, cw = emu(w / cols);
    const ln = (tag) => `<a:${tag} w="6350"><a:solidFill><a:srgbClr val="${PX.line}"/></a:solidFill></a:${tag}>`;
    const trs = rows.map((r, ri) => `<a:tr h="${emu(rowH[ri])}">${r.map(c => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${c ? para(pRun(c, { sz, b: ri === 0 }), { spaceAfter: 0 }) : '<a:p><a:endParaRPr lang="ar-SA"/></a:p>'}</a:txBody><a:tcPr marL="${emu(0.08)}" marR="${emu(0.08)}" marT="${emu(0.05)}" marB="${emu(0.05)}">${ln('lnL')}${ln('lnR')}${ln('lnT')}${ln('lnB')}${ri === 0 ? `<a:solidFill><a:srgbClr val="${PX.head}"/></a:solidFill>` : '<a:noFill/>'}</a:tcPr></a:tc>`).join('')}</a:tr>`).join('');
    return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="جدول ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(total)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr rtl="1" firstRow="1"/><a:tblGrid>${Array.from({ length: cols }, () => `<a:gridCol w="${cw}"/>`).join('')}</a:tblGrid>${trs}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}
function titleShapes(title, cont) {
    const t = cont ? `${title || ''} (تابع)` : (title || '');
    return textBox(2, 0.35, 0.95, [para(pRun(t, { sz: 30, b: true, color: PX.title }))], { name: 'العنوان' })
        + `<p:sp><p:nvSpPr><p:cNvPr id="3" name="خط"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(MARGIN)}" y="${emu(1.3)}"/><a:ext cx="${emu(CONTENT_W)}" cy="${emu(0.05)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${PX.orange}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>`;
}


// تصدير موضعي: يطابق مكان العناصر على لوحة الشريحة (نسب مئوية من مسرح 16:9)
const CQW = SLIDE_W * 0.72; // نقطة لكل 1cqw تقريبًا (≈9.6pt)
function positionedItems(slide, sizes) {
    const items = [];
    for (const b of slide.blocks) {
        const bx = b.box, X = bx.x / 100 * SLIDE_W, Y = bx.y / 100 * SLIDE_H, W = bx.w / 100 * SLIDE_W, H = Math.max(0.3, bx.h / 100 * SLIDE_H);
        const pt = (cqw) => Math.max(8, Math.round(cqw * CQW));
        if (b.type === 'text' && b.text.trim()) {
            const lines = b.text.split('\n').filter(l => l.trim());
            const ps = b.style === 'h' ? [para(pRun(b.text, { sz: pt(2.3), b: true, color: PX.orange }))] : lines.map(l => para(pRun(l, { sz: pt(1.7) }), { bullet: b.style === 'ul' }));
            items.push({ y: Y, sp: (id) => textBox(id, Y, H, ps, { x: X, w: W }) });
        } else if (b.type === 'finding' && `${b.text}${b.cause}${b.action}`.trim()) {
            const sev = SEVERITY_CHOICES.includes(b.severity) ? b.severity : severityFor(cleanText(`${b.text} ${b.cause}`), 'مشكلة');
            const col = sev === 'مرتفع' ? PX.high : sev === 'متوسط' ? PX.mid : PX.low, bg = sev === 'مرتفع' ? PX.highBg : sev === 'متوسط' ? PX.midBg : PX.lowBg;
            const meta = [`الخطورة: ${sev}`, `الحالة: ${STATUS_LABELS[b.status] || ''}`, b.owner && `المسؤول: ${b.owner}`, b.due && `الموعد: ${b.due}`].filter(Boolean).join('  ·  ');
            const ps = [para(pRun(meta, { sz: pt(1.3), b: true, color: col }), { spaceAfter: 3 })];
            if (b.text.trim()) ps.push(para(pRun(b.text, { sz: pt(1.7), b: true })));
            if (b.cause.trim()) ps.push(para(pRun('السبب المحتمل: ', { sz: pt(1.5), b: true, color: PX.muted }) + pRun(b.cause, { sz: pt(1.5) })));
            if (b.action.trim()) ps.push(para(pRun('الإجراء المقترح: ', { sz: pt(1.5), b: true, color: PX.low }) + pRun(b.action, { sz: pt(1.5) })));
            items.push({ y: Y, sp: (id) => textBox(id, Y, H, ps, { fill: bg, line: col, inset: 0.12, name: 'مؤشر', x: X, w: W }) });
        } else if (b.type === 'image' && sizes[b.assetId]) {
            const { w, h: ih } = sizes[b.assetId], capH = b.caption ? 0.36 : 0, availH = Math.max(0.2, H - capH);
            const k = Math.min(W / w, availH / ih), pw = w * k, ph = ih * k, px = X + (W - pw) / 2;
            items.push({ y: Y, sp: (id, _y, ridOf) => picture(id, ridOf(b.assetId), px, Y, pw, ph, b.caption) + (b.caption ? textBox(id + 1, Y + ph + 0.02, 0.34, [para(pRun(b.caption, { sz: 12, color: PX.muted }), { algn: 'ctr' })], { x: X, w: W }) : '') });
        } else if (b.type === 'table' && b.rows.length) {
            const rh = H / b.rows.length;
            items.push({ y: Y, sp: (id) => table(id, Y, b.rows, b.rows.map(() => rh), H, X, W, pt(1.3)) });
        } else if (b.type === 'comment' && b.text.trim()) {
            const t = `تعليق${b.author ? ` (${b.author})` : ''}: ${b.text}`;
            items.push({ y: Y, sp: (id) => textBox(id, Y, H, [para(pRun(t, { sz: pt(1.4), i: true, color: PX.muted }))], { fill: 'FFF6C8', line: 'E3D27A', inset: 0.1, x: X, w: W }) });
        }
    }
    return items;
}
const isPositioned = (slide) => slide.blocks.length > 0 && slide.blocks.every(b => b.box && isFinite(b.box.x) && isFinite(b.box.w));

async function imageSizes(blobs) {
    const out = {};
    for (const [id, blob] of Object.entries(blobs)) {
        if (!/^image\/(png|jpeg|jpg|gif)$/.test(blob.type)) continue;
        try { const bmp = await createImageBitmap(blob); out[id] = { w: bmp.width, h: bmp.height, type: blob.type === 'image/jpg' ? 'image/jpeg' : blob.type }; bmp.close?.(); } catch { /* صورة تالفة تُهمل */ }
    }
    return out;
}

async function buildPptx(project) {
    const blobs = await getAssetBlobs(project.id).catch(() => ({}));
    const sizes = await imageSizes(blobs);
    const media = []; // [{name, bytes}]
    const mediaOf = {};
    for (const id of Object.keys(sizes)) {
        const ext = sizes[id].type === 'image/png' ? 'png' : sizes[id].type === 'image/gif' ? 'gif' : 'jpeg';
        mediaOf[id] = `image${media.length + 1}.${ext}`;
        media.push({ name: mediaOf[id], blob: blobs[id] });
    }
    // توزيع العناصر على شرائح (تقسيم تلقائي عند الامتلاء)
    const pages = [];
    for (const slide of project.slides) {
        if (isPositioned(slide)) {
            const its = positionedItems(slide, sizes);
            if (slide.notes?.trim()) its.push({ y: 7.0, sp: (id) => textBox(id, 7.0, 0.4, [para(pRun(`ملاحظات المتحدث: ${slide.notes.trim().slice(0, 300)}`, { sz: 11, color: PX.muted }))]) });
            pages.push({ title: slide.title || 'بلا عنوان', cont: false, items: its, positioned: true });
            continue;
        }
        const items = buildSlideItems(slide, project, sizes);
        let page = { title: slide.title || 'بلا عنوان', cont: false, items: [] }, y = TOP;
        for (const it of items) {
            if (y + it.h > BOTTOM && page.items.length) { pages.push(page); page = { title: slide.title || 'بلا عنوان', cont: true, items: [] }; y = TOP; }
            page.items.push({ ...it, y }); y += it.h;
        }
        pages.push(page);
    }
    const files = [], slideRels = [];
    pages.forEach((page, pi) => {
        const rels = [];
        const ridOf = (assetId) => { let i = rels.findIndex(r => r.assetId === assetId); if (i < 0) { rels.push({ assetId }); i = rels.length - 1; } return `rId${i + 2}`; };
        let id = 10;
        const body = page.items.map(it => { const xml = it.sp(id, it.y, ridOf); id += 3; return xml; }).join('');
        files.push([`ppt/slides/slide${pi + 1}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${titleShapes(page.title, page.cont)}${body}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`]);
        files.push([`ppt/slides/_rels/slide${pi + 1}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>${rels.map((r, i) => `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaOf[r.assetId]}"/>`).join('')}</Relationships>`]);
    });
    const n = pages.length, R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    const exts = new Set(media.map(m => m.name.split('.').pop()));
    files.unshift(['[Content_Types].xml', `${head}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${[...exts].map(e => `<Default Extension="${e}" ContentType="image/${e}"/>`).join('')}<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>${pages.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}</Types>`]);
    files.push(['_rels/.rels', `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`]);
    files.push(['docProps/core.xml', `${head}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xe(project.name)}</dc:title><dc:creator>${xe(project.owner || 'وَرَق')}</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created></cp:coreProperties>`]);
    files.push(['ppt/presentation.xml', `${head}<p:presentation ${NS} rtl="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${pages.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 3}"/>`).join('')}</p:sldIdLst><p:sldSz cx="${emu(SLIDE_W)}" cy="${emu(SLIDE_H)}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`]);
    files.push(['ppt/_rels/presentation.xml.rels', `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${R}/theme" Target="theme/theme1.xml"/>${pages.map((_, i) => `<Relationship Id="rId${i + 3}" Type="${R}/slide" Target="slides/slide${i + 1}.xml"/>`).join('')}</Relationships>`]);
    const grp = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
    files.push(['ppt/slideMasters/slideMaster1.xml', `${head}<p:sldMaster ${NS}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree>${grp}</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr algn="r" rtl="1"><a:defRPr sz="3000"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr algn="r" rtl="1"><a:defRPr sz="1800"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr algn="r" rtl="1"><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>`]);
    files.push(['ppt/slideMasters/_rels/slideMaster1.xml.rels', `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${R}/theme" Target="../theme/theme1.xml"/></Relationships>`]);
    files.push(['ppt/slideLayouts/slideLayout1.xml', `${head}<p:sldLayout ${NS} type="blank" preserve="1"><p:cSld name="فارغ"><p:spTree>${grp}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`]);
    files.push(['ppt/slideLayouts/_rels/slideLayout1.xml.rels', `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`]);
    const fill3 = '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>';
    const ln3 = '<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="28575"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>';
    const eff3 = '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>';
    const font = (tag) => `<a:${tag}><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface="Arial"/></a:${tag}>`;
    const acc = ['E96826', '668C74', 'D9A441', 'AD5349', '5E8469', '29453C'].map((c, i) => `<a:accent${i + 1}><a:srgbClr val="${c}"/></a:accent${i + 1}>`).join('');
    files.push(['ppt/theme/theme1.xml', `${head}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="وَرَق"><a:themeElements><a:clrScheme name="Waraq"><a:dk1><a:srgbClr val="303532"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="29453C"/></a:dk2><a:lt2><a:srgbClr val="F4F1E9"/></a:lt2>${acc}<a:hlink><a:srgbClr val="E96826"/></a:hlink><a:folHlink><a:srgbClr val="6B6B62"/></a:folHlink></a:clrScheme><a:fontScheme name="Waraq">${font('majorFont')}${font('minorFont')}</a:fontScheme><a:fmtScheme name="Waraq">${fill3}${ln3}${eff3}<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`]);
    for (const m of media) files.push([`ppt/media/${m.name}`, new Uint8Array(await m.blob.arrayBuffer())]);
    return zipStore(files);
}
async function downloadPptx(project) {
    const blob = await buildPptx(project);
    downloadFile(`${project.name}.pptx`, blob, blob.type);
}

/* ================= الأيقونات (SVG مضمّنة) ================= */
/* ================= الأيقونات (SVG مضمّنة) ================= */
const ICON_PATHS = {
  plus: '<path d="M5 12h14"/><path d="M12 5v14"/>',
  x: '<path d="M18 6 6 18"/><path d="m6 6 12 12"/>',
  check: '<path d="M20 6 9 17l-5-5"/>',
  search: '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/>',
  chevronDown: '<path d="m6 9 6 6 6-6"/>',
  trash: '<path d="M3 6h18"/><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/><path d="M10 11v6"/><path d="M14 11v6"/>',
  pencil: '<path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z"/><path d="m15 5 4 4"/>',
  upload: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m17 8-5-5-5 5"/><path d="M12 3v12"/>',
  download: '<path d="M12 17V3"/><path d="m6 11 6 6 6-6"/><path d="M19 21H5"/>',
  shield: '<path d="M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1.17 1.17 0 0 1 1.52 0C14.51 3.81 17 5 19 5a1 1 0 0 1 1 1z"/><path d="m9 12 2 2 4-4"/>',
  fileText: '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M10 9H8"/><path d="M16 13H8"/><path d="M16 17H8"/>',
  fileJson: '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M10 12a1 1 0 0 0-1 1v1a1 1 0 0 1-1 1 1 1 0 0 1 1 1v1a1 1 0 0 0 1 1"/><path d="M14 18a1 1 0 0 0 1-1v-1a1 1 0 0 1 1-1 1 1 0 0 1-1-1v-1a1 1 0 0 0-1-1"/>',
  fileWarning: '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M12 9v4"/><path d="M12 17h.01"/>',
  presentation: '<path d="M2 3h20"/><path d="M21 3v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V3"/><path d="m7 21 5-5 5 5"/>',
  bookOpen: '<path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/>',
  hardDrive: '<path d="M22 12H2"/><path d="M5.45 5.11 2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/><path d="M6 16h.01"/><path d="M10 16h.01"/>',
  loader: '<path d="M21 12a9 9 0 1 1-6.219-8.56"/>',
  alert: '<circle cx="12" cy="12" r="10"/><path d="M12 8v4"/><path d="M12 16h.01"/>',
  stickyNote: '<path d="M16 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V8Z"/><path d="M15 3v4a2 2 0 0 0 2 2h4"/>',
  clock: '<circle cx="12" cy="12" r="10"/><path d="M12 6v6h4"/>',
  sliders: '<path d="M21 4h-7"/><path d="M10 4H3"/><path d="M21 12h-9"/><path d="M8 12H3"/><path d="M21 20h-5"/><path d="M12 20H3"/><path d="M14 2v4"/><path d="M8 10v4"/><path d="M16 18v4"/>',
  quote: '<path d="M16 3a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2 1 1 0 0 1 1 1v1a2 2 0 0 1-2 2 1 1 0 0 0-1 1v2a1 1 0 0 0 1 1 6 6 0 0 0 6-6V5a2 2 0 0 0-2-2z"/><path d="M5 3a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2 1 1 0 0 1 1 1v1a2 2 0 0 1-2 2 1 1 0 0 0-1 1v2a1 1 0 0 0 1 1 6 6 0 0 0 6-6V5a2 2 0 0 0-2-2z"/>',
  imagePlus: '<path d="M16 5h6"/><path d="M19 2v6"/><path d="M21 11.5V19a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h7.5"/><path d="m21 15-3.086-3.086a2 2 0 0 0-2.828 0L6 21"/><circle cx="9" cy="9" r="2"/>',
};
function icon(name, size = 16, extraClass = '') {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"${extraClass ? ` class="${extraClass}"` : ''}>${ICON_PATHS[name] || ''}</svg>`;
}

Object.assign(ICON_PATHS, {
  image: '<rect width="18" height="18" x="3" y="3" rx="2"/><circle cx="9" cy="9" r="2"/><path d="m21 15-3.086-3.086a2 2 0 0 0-2.828 0L6 21"/>',
  chevronLeft: '<path d="m15 18-6-6 6-6"/>',
  chevronRight: '<path d="m9 18 6-6-6-6"/>',
  table: '<path d="M12 3v18"/><rect width="18" height="18" x="3" y="3" rx="2"/><path d="M3 9h18"/><path d="M3 15h18"/>',
  flag: '<path d="M4 15s1-1 4-1 5 2 8 2 4-1 4-1V3s-1 1-4 1-5-2-8-2-4 1-4 1z"/><path d="M4 22v-7"/>',
  listChecks: '<path d="m3 17 2 2 4-4"/><path d="m3 7 2 2 4-4"/><path d="M13 6h8"/><path d="M13 12h8"/><path d="M13 18h8"/>',
  zoom: '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/><path d="M11 8v6"/><path d="M8 11h6"/>',
  layers: '<path d="m12 2 10 6-10 6L2 8z"/><path d="m2 12 10 6 10-6"/><path d="m2 16 10 6 10-6"/>',
  mic: '<path d="M12 2a3 3 0 0 0-3 3v7a3 3 0 0 0 6 0V5a3 3 0 0 0-3-3Z"/><path d="M19 10v2a7 7 0 0 1-14 0v-2"/><path d="M12 19v3"/>',
});

Object.assign(ICON_PATHS, {
  copy: '<rect width="14" height="14" x="8" y="8" rx="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/>',
  play: '<polygon points="6 3 20 12 6 21 6 3"/>',
  save: '<path d="M15.2 3a2 2 0 0 1 1.4.6l3.8 3.8a2 2 0 0 1 .6 1.4V19a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z"/><path d="M17 21v-7a1 1 0 0 0-1-1H8a1 1 0 0 0-1 1v7"/><path d="M7 3v4a1 1 0 0 0 1 1h7"/>',
  arrowUp: '<path d="m5 12 7-7 7 7"/><path d="M12 19V5"/>',
  arrowDown: '<path d="M12 5v14"/><path d="m19 12-7 7-7-7"/>',
  folder: '<path d="M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z"/>',
  message: '<path d="M7.9 20A9 9 0 1 0 4 16.1L2 22Z"/>',
  heading: '<path d="M6 12h12"/><path d="M6 20V4"/><path d="M18 20V4"/>',
  list: '<path d="M3 12h.01"/><path d="M3 18h.01"/><path d="M3 6h.01"/><path d="M8 12h13"/><path d="M8 18h13"/><path d="M8 6h13"/>',
  type: '<polyline points="4 7 4 4 20 4 20 7"/><line x1="9" x2="15" y1="20" y2="20"/><line x1="12" x2="12" y1="4" y2="20"/>',
  user: '<path d="M19 21v-2a4 4 0 0 0-4-4H9a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>',
  calendar: '<path d="M8 2v4"/><path d="M16 2v4"/><rect width="18" height="18" x="3" y="4" rx="2"/><path d="M3 10h18"/>',
  menu: '<path d="M4 12h16"/><path d="M4 6h16"/><path d="M4 18h16"/>',
  undo: '<path d="M3 7v6h6"/><path d="M21 17a9 9 0 0 0-9-9 9 9 0 0 0-6 2.3L3 13"/>',
  eye: '<path d="M2.06 12.35a1 1 0 0 1 0-.7 10.75 10.75 0 0 1 19.88 0 1 1 0 0 1 0 .7 10.75 10.75 0 0 1-19.88 0"/><circle cx="12" cy="12" r="3"/>',
  grip: '<circle cx="9" cy="6" r="1"/><circle cx="9" cy="12" r="1"/><circle cx="9" cy="18" r="1"/><circle cx="15" cy="6" r="1"/><circle cx="15" cy="12" r="1"/><circle cx="15" cy="18" r="1"/>',
  barChart: '<path d="M3 3v16a2 2 0 0 0 2 2h16"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/>',
});

Object.assign(ICON_PATHS, {
  redo: '<path d="M21 7v6h-6"/><path d="M3 17a9 9 0 0 1 9-9 9 9 0 0 1 6 2.3l3 2.7"/>',
  zoomIn: '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/><path d="M11 8v6"/><path d="M8 11h6"/>',
  zoomOut: '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/><path d="M8 11h6"/>',
  pen: '<path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/>',
  arrowUpRight: '<path d="M7 7h10v10"/><path d="M7 17 17 7"/>',
  square: '<rect width="18" height="18" x="3" y="3" rx="2"/>',
  circle: '<circle cx="12" cy="12" r="10"/>',
  highlighter: '<path d="m9 11-6 6v3h9l3-3"/><path d="m22 12-4.6 4.6a2 2 0 0 1-2.8 0l-5.2-5.2a2 2 0 0 1 0-2.8L14 4"/>',
  hash: '<line x1="4" x2="20" y1="9" y2="9"/><line x1="4" x2="20" y1="15" y2="15"/><line x1="10" x2="8" y1="3" y2="21"/><line x1="16" x2="14" y1="3" y2="21"/>',
  crop: '<path d="M6 2v14a2 2 0 0 0 2 2h14"/><path d="M18 22V8a2 2 0 0 0-2-2H2"/>',
  pointer: '<path d="M4.037 4.688a.495.495 0 0 1 .651-.651l16 6.5a.5.5 0 0 1-.063.947l-6.124 1.58a2 2 0 0 0-1.438 1.435l-1.579 6.126a.5.5 0 0 1-.947.063z"/>',
  front: '<rect width="8" height="8" x="8" y="8" rx="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/>',
  alignH: '<path d="M2 12h20"/><rect width="6" height="8" x="4" y="4" rx="1"/><rect width="6" height="6" x="14" y="9" rx="1"/>',
  alignV: '<path d="M12 2v20"/><rect width="8" height="6" x="4" y="4" rx="1"/><rect width="6" height="6" x="9" y="14" rx="1"/>',
  maximize: '<path d="M8 3H5a2 2 0 0 0-2 2v3"/><path d="M21 8V5a2 2 0 0 0-2-2h-3"/><path d="M3 16v3a2 2 0 0 0 2 2h3"/><path d="M16 21h3a2 2 0 0 0 2-2v-3"/>',
  rotateCcw: '<path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/>',
  layoutList: '<rect width="7" height="7" x="3" y="3" rx="1"/><rect width="7" height="7" x="3" y="14" rx="1"/><path d="M14 4h7"/><path d="M14 9h7"/><path d="M14 15h7"/><path d="M14 20h7"/>',
  monitor: '<rect width="20" height="14" x="2" y="3" rx="2"/><path d="M8 21h8"/><path d="M12 17v4"/>',
});

/* ================= الحالة ================= */
const OPTION_KEY = 'waraq.allImages';
function readOption() { try { return localStorage.getItem(OPTION_KEY) === '1'; } catch { return false; } }
function writeOption(value) { try { localStorage.setItem(OPTION_KEY, value ? '1' : '0'); } catch { /* التخزين غير متاح */ } }

const state = {
  projects: [], projectId: '', mode: 'edit', slideIdx: 0, loadingSaved: true, progress: null, error: '', notice: '',
  allImages: readOption(), saveState: 'saved', sidebarOpen: false, projectSearch: '', present: null,
  tab: 'cases', severity: 'الكل', sectionFilter: 'الكل', statusFilter: 'الكل', slideFilter: '', search: '',
  sideCollapsed: false, editView: 'stage', sel: null, zoom: 1, editingId: null, hist: { past: [], future: [] }, annot: null,
  ownerFilter: 'الكل', view: 'project', boardOwner: 'الكل', boardScope: 'all', boardProject: 'الكل', drag: null,
  sortBy: 'severity', groupBy: 'none', slidesFilter: 'all',
  selectedFinding: null, selectedSlide: null, lightbox: null,
  projectForm: null, deleteProject: null, deleteSlide: null, undo: null, focusId: null, exportOpen: false,
  assetUrls: {},
};
const $app = document.getElementById('app');
const $fileInput = document.getElementById('file-input');
const $imageInput = document.getElementById('image-input');
const h = escapeHtml;
const currentProject = () => state.projects.find(p => p.id === state.projectId) || null;
/** التحليل الحي للمشروع المفتوح (تُعاد قراءته فقط عند تغيّر المشروع). */
const active = () => { const p = currentProject(); return p ? analysisOf(p).analysis : null; };
const activeSummary = () => { const p = currentProject(); return p ? analysisOf(p).summary : null; };
const pad2 = (n) => String(n).padStart(2, '0');

/* ---------- الحفظ التلقائي ---------- */
let saveTimer = 0, liveTimer = 0;
function setSaveIndicator() {
  const el = document.getElementById('save-indicator');
  if (!el) return;
  const map = { saved: ['تم الحفظ', 'ok'], saving: ['جارٍ الحفظ…', 'busy'], error: ['تعذّر الحفظ', 'bad'], dirty: ['تغييرات غير محفوظة', 'busy'] };
  const [label, tone] = map[state.saveState] || map.saved;
  el.className = `save-indicator save-${tone}`;
  el.innerHTML = `${icon(tone === 'ok' ? 'check' : tone === 'bad' ? 'alert' : 'loader', 13, tone === 'busy' ? 'spin' : '')}<span>${label}</span>`;
}
async function persist(project) {
  state.saveState = 'saving'; setSaveIndicator();
  try { await saveProject(project); state.saveState = 'saved'; queueSync(); }
  catch { state.saveState = 'error'; }
  setSaveIndicator();
}
function scheduleSave(project) {
  state.saveState = 'dirty'; setSaveIndicator();
  clearTimeout(saveTimer);
  saveTimer = window.setTimeout(() => { saveTimer = 0; void persist(project); }, 600);
}
async function flushSave() {
  if (!saveTimer) return;
  clearTimeout(saveTimer); saveTimer = 0;
  const p = currentProject();
  if (p) await persist(p);
}
/** أي تعديل على المشروع يمر من هنا: يحدّث التاريخ ويُبطل التحليل ويجدول الحفظ. */
function touch(project) {
  project.updatedAt = new Date().toISOString();
  project._rev = (project._rev || 0) + 1;
  scheduleSave(project);
  clearTimeout(liveTimer);
  liveTimer = window.setTimeout(refreshLive, 250);
}
/** تحديث الأرقام الحية (شريط التحليل المباشر وقائمة المشاريع) دون إعادة رسم المحرّر. */
function refreshLive() {
  const p = currentProject();
  if (!p) return;
  const s = analysisOf(p).summary;
  const set = (id, text) => { const el = document.getElementById(id); if (el) el.textContent = text; };
  set('live-cases', String(s.cases.length));
  set('live-high', String(s.bySeverity['مرتفع']));
  set('live-open', String(s.cases.length - s.byStatus.done));
  const risk = document.getElementById('live-risk');
  if (risk) { risk.textContent = s.risk; risk.className = `risk-badge ${RISK_TONE[s.risk]}`; }
  const list = document.getElementById('project-list');
  if (list) list.innerHTML = projectListHtml();
  const an = document.getElementById('analysis-list');
  if (an) an.innerHTML = analysisListHtml();
  const strip = document.getElementById('slide-strip-list');
  if (strip) strip.innerHTML = slideStripItems(p);
}

/* ----- مساعدات الصور ----- */
/* ================= الصور (object URLs للتحليل المفتوح فقط) ================= */
let assetLoadToken = 0;
function releaseAssetUrls() { Object.values(state.assetUrls).forEach(url => URL.revokeObjectURL(url)); state.assetUrls = {}; }
async function loadAssetUrls(analysisId) {
  const token = ++assetLoadToken;
  releaseAssetUrls();
  if (!analysisId) return;
  try {
    const blobs = await getAssetBlobs(analysisId);
    if (token !== assetLoadToken) return;
    const proj = state.projects.find(x => x.id === analysisId);
    const used = new Set(proj ? proj.slides.flatMap(sl => sl.blocks.flatMap(b => [b.assetId, b.srcAssetId])).filter(Boolean) : Object.keys(blobs));
    for (const [id, blob] of Object.entries(blobs)) {
      if (!used.has(id)) { void deleteAssetBlob(analysisId, id).catch(() => {}); continue; } // صور يتيمة (بعد التعليم/القص/حذف العناصر)
      state.assetUrls[id] = URL.createObjectURL(blob);
    }
    render();
  } catch { /* تظهر رسالة بديلة بدل الصورة */ }
}
const imageSrc = (a, id) => state.assetUrls[id] || a?.imageAssets?.[id]?.dataUrl;
/** كل صور العرض بترتيب الشرائح (للمعرض والـ lightbox). */
function galleryItems(a) {
  return a.slides.flatMap(slide => (slide.imageIds || []).map(id => ({ slide, id, asset: a.imageAssets?.[id] })))
    .filter(item => item.asset && !item.asset.error && !item.asset.externalUrl);
}


/* ----- الملخص التنفيذي ----- */
const RISK_TONE = { مرتفع: 'risk-high', متوسط: 'risk-mid', منخفض: 'risk-low' };
function summaryHtml(a, s) {
  const total = s.cases.length;
  const pct = (n) => total ? n / total * 100 : 0;
  const withCases = s.slides.length;
  const st = a.imageStats;
  const imagesText = st ? `${st.kept}` : String(galleryItems(a).length);
  const imagesNote = st && st.skippedTotal ? `من ${st.found} · استُبعد ${st.skippedTotal}` : 'في المشروع';
  const headline = total
    ? `${total} حالة في ${withCases} من ${a.slideCount} شريحة`
    : 'لم تُرصد حالات في النص المستخرج';
  const sub = total
    ? (s.openHigh ? `${s.openHigh} حالة مرتفعة الأهمية بانتظار المتابعة؛ ابدأ بها من قائمة الأولويات.` : (s.byStatus.done === total ? 'أُغلقت كل الحالات.' : 'لا توجد حالات مرتفعة الأهمية مفتوحة.'))
    : 'أضف نصوصًا أو مؤشرات في المحرّر ليحلّلها النظام فورًا.';
  return `<section class="summary-band animate-in delay-1">
    <div class="summary-main">
      <span class="risk-badge ${RISK_TONE[s.risk]}">${icon('flag', 14)}مستوى الانتباه: ${s.risk}</span>
      <h2>${h(headline)}</h2><p>${h(sub)}</p>
      ${total ? `<div class="progress-stack" role="img" aria-label="تقدم المتابعة ${s.progress}%"><i class="ps-done" style="width:${pct(s.byStatus.done)}%"></i><i class="ps-progress" style="width:${pct(s.byStatus.progress)}%"></i></div>
      <div class="progress-legend"><span><i class="ps-done"></i>مغلقة ${s.byStatus.done}</span><span><i class="ps-progress"></i>قيد المتابعة ${s.byStatus.progress}</span><span><i class="ps-open"></i>مفتوحة ${s.byStatus.open}</span><strong>${s.progress}% مُنجز</strong></div>` : ''}
    </div>
    <div class="summary-stats">
      <div class="mini-stat"><span>الشرائح</span><strong>${a.slideCount}</strong><small>${a.slides.filter(x => x.text.trim()).length} فيها نص</small>${icon('presentation', 17)}</div>
      <div class="mini-stat"><span>الحالات</span><strong>${total}</strong><small>${s.bySeverity['مرتفع']} مرتفعة · ${s.bySeverity['متوسط']} متوسطة</small>${icon('fileWarning', 17)}</div>
      <div class="mini-stat"><span>الصور</span><strong>${imagesText}</strong><small>${h(imagesNote)}</small>${icon('image', 17)}</div>
      <div class="mini-stat"><span>التعليقات</span><strong>${a.notes.length}</strong><small>على الشرائح</small>${icon('stickyNote', 17)}</div>
    </div></section>`;
}

function insightsHtml(a, s) {
  const total = s.cases.length;
  const sevRows = ['مرتفع', 'متوسط', 'منخفض'].map((name, index) => {
    const count = s.bySeverity[name];
    return `<button class="chart-line ${state.severity === name ? 'chart-selected' : ''}" data-action="chart-severity" data-value="${name}">
      <span class="chart-label"><i class="dot ${severityTone[name]}"></i>${name}</span>
      <span class="bar-track"><span class="bar-fill bar-${index}" style="width:${total ? Math.max(count ? 8 : 0, count / total * 100) : 0}%"></span></span><strong>${count}</strong></button>`;
  }).join('');
  const topSections = s.sections.slice(0, 5), maxSection = topSections[0]?.[1] || 1;
  const catRows = topSections.length ? topSections.map(([name, count], index) => `
    <button class="chart-line category-chart-line ${state.sectionFilter === name ? 'chart-selected' : ''}" data-action="chart-section" data-value="${h(name)}">
      <span class="chart-label"><i class="dot category-dot category-dot-${index}"></i>${h(name)}</span>
      <span class="bar-track"><span class="bar-fill category-bar-${index}" style="width:${Math.max(8, count / maxSection * 100)}%"></span></span><strong>${count}</strong></button>`).join('')
    : '<p class="category-empty">لا توجد حالات موزعة على أقسام بعد.</p>';
  const topSlides = s.slides.slice(0, 5), maxSlide = topSlides[0]?.[1] || 1;
  const slideRows = topSlides.length ? topSlides.map(([number, count], index) => {
    const slide = a.slides.find(x => x.number === number);
    return `<button class="chart-line" data-action="goto-slide" data-id="${number}">
      <span class="chart-label"><i class="dot category-dot category-dot-${index}"></i>شريحة ${number}</span>
      <span class="bar-track"><span class="bar-fill category-bar-${index}" style="width:${Math.max(8, count / maxSlide * 100)}%"></span></span><strong>${count}</strong></button>
      <small class="chart-sub">${h((slide?.title || '').slice(0, 46))}</small>`;
  }).join('') : '<p class="category-empty">لا توجد شرائح بها حالات.</p>';
  const priorities = s.priorities.length ? s.priorities.map(c => `<button class="priority-item" data-action="goto-case" data-id="${h(c.id)}">
      <span class="severity-tag ${severityTone[c.severity]}">${c.severity}</span>
      <span class="priority-text"><strong>${h(c.title.slice(0, 90))}${c.title.length > 90 ? '…' : ''}</strong><small>شريحة ${c.slideNumber} · ${h(c.section)}</small></span>${icon('chevronLeft', 15)}</button>`).join('')
    : '<p class="category-empty">لا توجد حالات مفتوحة. أحسنت!</p>';
  return `<section class="insight-row animate-in delay-2">
    <div class="panel"><div class="panel-heading"><div><span class="section-kicker">الأولويات</span><h2>ابدأ من هنا</h2></div>${icon('listChecks', 18, 'heading-glyph')}</div><div class="priority-list">${priorities}</div></div>
    <div class="panel chart-panel"><div class="panel-heading"><div><span class="section-kicker">التوزيع</span><h2>الأهمية والأقسام</h2></div>${icon('sliders', 18, 'heading-glyph')}</div>
      <div class="severity-chart">${sevRows}</div><div class="chart-divider"></div><div class="severity-chart category-chart">${catRows}</div><div class="chart-footnote">اضغط على أي سطر لتصفية الحالات.</div></div>
    <div class="panel chart-panel"><div class="panel-heading"><div><span class="section-kicker">الشرائح</span><h2>الأكثر حالات</h2></div>${icon('layers', 18, 'heading-glyph')}</div>
      <div class="severity-chart category-chart">${slideRows}</div></div>
  </section>`;
}

/* ----- الصور المصغرة والمعرض ----- */
function thumbStripHtml(a, slide, limit = 3) {
  const ids = (slide.imageIds || []).filter(id => imageSrc(a, id));
  if (!ids.length) return '';
  const index = galleryItems(a);
  return `<div class="thumb-strip">${ids.slice(0, limit).map(id => {
    const pos = index.findIndex(item => item.id === id && item.slide.number === slide.number);
    return `<button class="thumb" data-action="lightbox" data-index="${pos}" aria-label="تكبير صورة الشريحة ${slide.number}"><img src="${h(imageSrc(a, id))}" alt="صورة الشريحة ${slide.number}" loading="lazy"></button>`;
  }).join('')}${ids.length > limit ? `<span class="thumb-more">+${ids.length - limit}</span>` : ''}</div>`;
}

function galleryHtml(slide) {
  const a = active();
  const ids = slide.imageIds || [];
  if (!ids.length) return '';
  const items = ids.map((imageId, index) => {
    const asset = a?.imageAssets?.[imageId];
    const src = imageSrc(a, imageId);
    const label = asset?.fileName || `صورة ${index + 1}`;
    const alt = `صورة الشريحة ${slide.number}: ${label}`;
    const ext = asset?.externalUrl;
    const safeExt = ext && /^https?:\/\//i.test(ext) ? ext : '';
    const body = src && !asset?.error
      ? `<a class="slide-image-link" href="${h(src)}" target="_blank" rel="noreferrer" aria-label="فتح ${h(alt)} بالحجم الأصلي"><img src="${h(src)}" alt="${h(alt)}" loading="lazy" decoding="async"></a>`
      : `<div class="slide-image-unavailable" role="img" aria-label="${h(alt)}"><span>${h(asset?.error || (ext ? 'هذه الصورة مرتبطة بمصدر خارجي وليست مضمّنة في الملف.' : 'جارٍ تحميل الصورة…'))}</span>${safeExt ? `<a href="${h(safeExt)}" target="_blank" rel="noreferrer">فتح المصدر الخارجي</a>` : ''}</div>`;
    return `<figure class="slide-image-card">${body}<figcaption>${h(label)}${asset?.contentType ? `<small>${h(asset.contentType)}</small>` : ''}</figcaption></figure>`;
  }).join('');
  return `<div class="slide-image-gallery" aria-label="صور الشريحة ${slide.number}">${items}</div>`;
}

/* ----- الحالات ----- */
function casesView(a, s) {
  const q = state.search.trim().toLocaleLowerCase('ar');
  let list = s.cases.filter(c =>
    (state.severity === 'الكل' || c.severity === state.severity) &&
    (state.sectionFilter === 'الكل' || c.section === state.sectionFilter) &&
    (state.statusFilter === 'الكل' || c.status === state.statusFilter) &&
    (state.ownerFilter === 'الكل' || (c.owner || 'غير مُسنَد') === state.ownerFilter) &&
    (!state.slideFilter || c.slideNumber === Number(state.slideFilter)) &&
    (!q || `${c.title} ${c.cause?.text ?? ''} ${c.action?.text ?? ''} ${c.section}`.toLocaleLowerCase('ar').includes(q)));
  const bySev = (x, y) => SEVERITY_RANK[x.severity] - SEVERITY_RANK[y.severity];
  list = list.slice().sort(state.sortBy === 'slide' ? (x, y) => x.slideNumber - y.slideNumber || bySev(x, y) : (x, y) => bySev(x, y) || x.slideNumber - y.slideNumber);
  return list;
}

function caseCardHtml(a, c) {
  const slide = a.slides.find(x => x.number === c.slideNumber);
  return `<article class="case-card ${severityTone[c.severity]} status-${c.status}">
    <div class="case-head"><div class="finding-meta"><span class="severity-tag ${severityTone[c.severity]}">${c.severity}</span><span class="finding-category">${h(c.section)}</span><span class="slide-tag">شريحة ${c.slideNumber}</span>${c.manual ? '<span class="slide-tag manual-tag">أدخلته يدويًا</span>' : ''}${c.owner ? `<span class="slide-tag">${icon('user', 11)}${h(c.owner)}</span>` : ''}${c.due ? `<span class="slide-tag">${icon('calendar', 11)}${h(c.due)}</span>` : ''}</div>
      <label class="status-select" title="حالة المتابعة"><select class="status-pick" data-case="${h(c.id)}" aria-label="حالة المتابعة">${Object.entries(STATUS_LABELS).map(([k, v]) => `<option value="${k}"${c.status === k ? ' selected' : ''}>${v}</option>`).join('')}</select></label></div>
    <h3 class="case-title">${h((c.observation ?? c.cause ?? c.action).text)}</h3>
    ${c.cause && c.observation ? `<div class="case-part cause"><span>السبب المحتمل</span><p>${h(c.cause.text)}</p></div>` : ''}
    ${c.action && (c.observation || c.cause) ? `<div class="case-part action"><span>الإجراء المقترح</span><p>${h(c.action.text)}</p></div>` : ''}
    <div class="case-foot">${slide ? thumbStripHtml(a, slide, 3) : ''}<div class="case-buttons">
      <button class="text-button" data-action="open-case" data-id="${h(c.id)}">${icon('quote', 14)}نص الشريحة</button>
      <button class="text-button" data-action="edit-case" data-id="${h(c.id)}">${icon('pencil', 14)}تحرير</button></div></div></article>`;
}

function casesListHtml(a, s) {
  const list = casesView(a, s);
  if (!list.length) {
    const any = s.cases.length > 0;
    return `<div class="empty-results"><div class="empty-results-icon">${icon('search', 20)}</div><h3>${any ? 'لا نتائج بهذه التصفية' : 'لا توجد حالات نصية'}</h3><p>${any ? 'جرّب تغيير الأهمية أو القسم أو الحالة أو رقم الشريحة أو عبارة البحث.' : 'لم يرصد النظام كلمات أو أنماطًا تستحق الإشارة. أضف «مؤشرًا» يدويًا في المحرّر بمستوى الخطورة الذي تحدده.'}</p>${any ? '<button class="text-button" data-action="clear-filters">إزالة التصفية</button>' : ''}</div>`;
  }
  if (state.groupBy === 'none') return `<div class="case-list">${list.map(c => caseCardHtml(a, c)).join('')}</div>`;
  const groups = new Map();
  for (const c of list) {
    const key = state.groupBy === 'section' ? c.section : c.slideNumber;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(c);
  }
  const entries = Array.from(groups.entries());
  if (state.groupBy === 'slide') entries.sort((x, y) => x[0] - y[0]);
  return entries.map(([key, items]) => {
    const slide = state.groupBy === 'slide' ? a.slides.find(x => x.number === key) : null;
    return `<section class="case-group"><h3 class="group-heading">${state.groupBy === 'slide' ? `شريحة ${key}${slide ? ` · ${h(slide.title.slice(0, 50))}` : ''}` : h(key)}<span>${items.length}</span></h3>
      <div class="case-list">${items.map(c => caseCardHtml(a, c)).join('')}</div></section>`;
  }).join('');
}

function casesTabHtml(a, s) {
  const opts = (list, current) => ['الكل', ...list].map(v => `<option value="${h(v)}"${v === current ? ' selected' : ''}>${h(v)}</option>`).join('');
  const sections = s.sections.map(x => x[0]).sort((x, y) => x.localeCompare(y, 'ar'));
  const statusOpts = `<option value="الكل"${state.statusFilter === 'الكل' ? ' selected' : ''}>الكل</option>` + Object.entries(STATUS_LABELS).map(([k, v]) => `<option value="${k}"${state.statusFilter === k ? ' selected' : ''}>${v}</option>`).join('');
  const ownerCounts = new Map();
  s.cases.filter(c => c.status !== 'done').forEach(c => { const o = c.owner || 'غير مُسنَد'; ownerCounts.set(o, (ownerCounts.get(o) ?? 0) + 1); });
  const ownerNames = Array.from(new Set(s.cases.map(c => c.owner || 'غير مُسنَد'))).sort((x, y) => x.localeCompare(y, 'ar'));
  const chips = ownerCounts.size > 1 || (ownerCounts.size === 1 && !ownerCounts.has('غير مُسنَد')) ? `<div class="owner-chips" aria-label="ما على كل مسؤول"><span>المسؤولون:</span>${Array.from(ownerCounts.entries()).map(([o, n]) => `<button class="owner-chip${state.ownerFilter === o ? ' on' : ''}" data-action="filter-owner" data-owner="${h(o)}">${icon('user', 11)}${h(o)} <b>${n}</b></button>`).join('')}</div>` : '';
  return chips + `<div class="filters">
      <label class="search-field">${icon('search')}<input id="search" value="${h(state.search)}" placeholder="ابحث في الحالات..." aria-label="ابحث في الحالات"></label>
      <label class="filter-select"><span>الأهمية</span><select id="severity" aria-label="تصفية حسب الأهمية">${opts(['مرتفع', 'متوسط', 'منخفض'], state.severity)}</select>${icon('chevronDown', 13)}</label>
      <label class="filter-select"><span>القسم</span><select id="section" aria-label="تصفية حسب القسم">${opts(sections, state.sectionFilter)}</select>${icon('chevronDown', 13)}</label>
      <label class="filter-select"><span>المتابعة</span><select id="status-filter" aria-label="تصفية حسب المتابعة">${statusOpts}</select>${icon('chevronDown', 13)}</label>
      <label class="filter-select"><span>المسؤول</span><select id="owner-filter" aria-label="تصفية حسب المسؤول">${opts(ownerNames, state.ownerFilter)}</select>${icon('chevronDown', 13)}</label>
      <label class="slide-filter"><span>الشريحة</span><input id="slide-filter" type="number" min="1" max="${a.slideCount}" placeholder="الكل" value="${h(state.slideFilter)}" aria-label="تصفية برقم الشريحة"></label>
    </div>
    <div class="view-options"><label class="filter-select"><span>الترتيب</span><select id="sort-by" aria-label="الترتيب"><option value="severity"${state.sortBy === 'severity' ? ' selected' : ''}>الأهمية</option><option value="slide"${state.sortBy === 'slide' ? ' selected' : ''}>رقم الشريحة</option></select>${icon('chevronDown', 13)}</label>
      <label class="filter-select"><span>التجميع</span><select id="group-by" aria-label="التجميع"><option value="none"${state.groupBy === 'none' ? ' selected' : ''}>بدون</option><option value="section"${state.groupBy === 'section' ? ' selected' : ''}>حسب القسم</option><option value="slide"${state.groupBy === 'slide' ? ' selected' : ''}>حسب الشريحة</option></select>${icon('chevronDown', 13)}</label></div>
    <div id="cases-list">${casesListHtml(a, s)}</div>`;
}

/* ----- الشرائح والصور ----- */
function slidesTabHtml(a, s) {
  const counts = new Map(s.slides);
  const highBySlide = new Map();
  s.cases.filter(c => c.severity === 'مرتفع').forEach(c => highBySlide.set(c.slideNumber, (highBySlide.get(c.slideNumber) ?? 0) + 1));
  const list = a.slides.filter(x => state.slidesFilter === 'all' || (state.slidesFilter === 'with') === counts.has(x.number));
  const filter = `<div class="filters"><label class="filter-select"><span>عرض</span><select id="slides-filter" aria-label="تصفية الشرائح"><option value="all"${state.slidesFilter === 'all' ? ' selected' : ''}>كل الشرائح</option><option value="with"${state.slidesFilter === 'with' ? ' selected' : ''}>التي بها حالات</option><option value="without"${state.slidesFilter === 'without' ? ' selected' : ''}>بلا حالات</option></select>${icon('chevronDown', 13)}</label></div>`;
  return filter + `<div class="slide-list">${list.map(slide => {
    const n = counts.get(slide.number) || 0, hi = highBySlide.get(slide.number) || 0;
    return `<article class="slide-record"><button class="slide-row" data-action="open-slide" data-id="${slide.number}"><span class="slide-number">${String(slide.number).padStart(2, '0')}</span>
      <span class="slide-content"><strong>${h(slide.title)}</strong><small>${h(slide.text.slice(0, 150) || 'لم يُستخرج نص من هذه الشريحة')}${slide.text.length > 150 ? '…' : ''}</small>
        <span class="slide-badges">${n ? `<span class="badge-case${hi ? ' hot' : ''}">${n} حالة${hi ? ` · ${hi} مرتفعة` : ''}</span>` : ''}${slide.notes ? `<span class="badge-soft">${icon('mic', 11)}ملاحظات متحدث</span>` : ''}${slide.tables?.length ? `<span class="badge-soft">${icon('table', 11)}جدول</span>` : ''}${slide.imageIds?.length ? `<span class="badge-soft">${icon('image', 11)}${slide.imageIds.length} صورة</span>` : ''}</span></span>
      <span class="slide-open-label">فتح النص ${icon('chevronDown', 14)}</span></button>${thumbStripHtml(a, slide, 4)}</article>`;
  }).join('') || '<div class="empty-results"><p>لا توجد شرائح بهذه التصفية.</p></div>'}</div>`;
}

function imagesTabHtml(a) {
  const items = galleryItems(a);
  const st = a.imageStats;
  const reasons = st && st.skippedTotal ? Object.entries(st.skipped).map(([k, n]) => `<span class="badge-soft">${h(SKIP_REASONS[k] || k)} · ${n}</span>`).join('') : '';
  const info = st
    ? `<div class="image-info"><strong>${st.mode === 'all' ? 'كل صور الملف المستورد' : 'الصور الأساسية من الملف المستورد'}</strong><span>اعتُمد ${st.kept} من ${st.found}${st.skippedTotal ? ` · استُبعد ${st.skippedTotal}` : ''}</span><div class="image-reasons">${reasons}</div>
      ${st.mode === 'main' && st.skippedTotal ? '<small>استُبعدت الشعارات والخلفيات والأيقونات والصور المتكررة عند الاستيراد. لاستيراد كل شيء فعّل «كل الصور» قبل الاستيراد.</small>' : ''}</div>`
    : '';
  if (!items.length) return info + `<div class="empty-results"><div class="empty-results-icon">${icon('image', 20)}</div><h3>لا توجد صور</h3><p>أضف صورًا من المحرّر (زر «صورة» أو لصقها مباشرة).</p></div>`;
  return info + `<div class="gallery-grid">${items.map((item, i) => `<figure class="gallery-card"><button data-action="lightbox" data-index="${i}" aria-label="تكبير الصورة">
    <img src="${h(imageSrc(a, item.id) || '')}" alt="صورة الشريحة ${item.slide.number}" loading="lazy">${icon('zoom', 16, 'zoom-icon')}</button>
    <figcaption><strong>شريحة ${item.slide.number}</strong><small>${h(item.slide.title.slice(0, 40))}</small></figcaption></figure>`).join('')}</div>`;
}

/* ================= محرّر الشرائح ================= */
const TEXT_STYLES = { p: 'فقرة', h: 'عنوان فرعي', ul: 'قائمة نقطية' };
const STATUS_OPTIONS = Object.entries(STATUS_LABELS);
const fid = (block, field) => `f-${block.id}-${field}`;
const selectHtml = (id, attrs, options, current) => `<select id="${id}" ${attrs}>${options.map(([v, l]) => `<option value="${h(v)}"${String(v) === String(current) ? ' selected' : ''}>${h(l)}</option>`).join('')}</select>`;

function slideBadgeCounts(slide) {
  let findings = 0, images = 0, comments = 0;
  for (const b of slide.blocks) { if (b.type === 'finding') findings++; else if (b.type === 'image') images++; else if (b.type === 'comment') comments++; }
  return { findings, images, comments };
}
function slideStripItems(p) {
  const s = analysisOf(p).summary;
  const perSlide = new Map(s.slides);
  return p.slides.map((slide, i) => {
    const c = slideBadgeCounts(slide);
    const n = perSlide.get(i + 1) || 0;
    return `<button class="strip-item${i === state.slideIdx ? ' strip-active' : ''}" draggable="true" data-drag-kind="slide" data-drag-id="${i}" data-action="select-slide" data-index="${i}" aria-label="الشريحة ${i + 1}">
      <span class="strip-num">${pad2(i + 1)}</span><span class="strip-body">${miniHtml(slide)}<strong>${h(slide.title || 'بلا عنوان')}</strong>
      <small>${n ? `<span class="strip-dot${s.cases.some(x => x.slideNumber === i + 1 && x.severity === 'مرتفع') ? ' hot' : ''}"></span>${n} حالة` : 'بلا حالات'}${c.images ? ` · ${c.images} صورة` : ''}${c.comments ? ` · ${c.comments} تعليق` : ''}</small></span></button>`;
  }).join('');
}

function blockTools(b, i, count) {
  return `<div class="block-tools">
    <button data-action="move-block" data-id="${b.id}" data-dir="-1" aria-label="تحريك للأعلى" title="تحريك للأعلى"${i === 0 ? ' disabled' : ''}>${icon('arrowUp', 14)}</button>
    <button data-action="move-block" data-id="${b.id}" data-dir="1" aria-label="تحريك للأسفل" title="تحريك للأسفل"${i === count - 1 ? ' disabled' : ''}>${icon('arrowDown', 14)}</button>
    <button data-action="dup-block" data-id="${b.id}" aria-label="نسخ العنصر" title="نسخ">${icon('copy', 14)}</button>
    <button data-action="del-block" data-id="${b.id}" aria-label="حذف العنصر" title="حذف" class="danger">${icon('trash', 14)}</button></div>`;
}

function blockHtml(b, i, count) {
  const head = (label, ico, extra = '') => `<div class="block-head"><span class="grip" draggable="true" data-drag-kind="block" data-drag-id="${b.id}" title="اسحب لإعادة الترتيب أو لنقله إلى شريحة أخرى" aria-label="اسحب لإعادة الترتيب">${icon('grip', 15)}</span><span class="block-type">${icon(ico, 14)}${label}</span>${extra}${blockTools(b, i, count)}</div>`;
  const f = (field) => `data-block="${b.id}" data-field="${field}"`;
  if (b.type === 'text') {
    return `<section class="block block-text" id="block-${b.id}">${head('نص', 'type', selectHtml(fid(b, 'style'), f('style') + ' class="mini-select" aria-label="نمط النص"', Object.entries(TEXT_STYLES), b.style))}
      <textarea class="autosize block-area style-${b.style}" id="${fid(b, 'text')}" ${f('text')} rows="2" placeholder="${b.style === 'ul' ? 'كل سطر نقطة مستقلة…' : 'اكتب النص هنا…'}">${h(b.text)}</textarea></section>`;
  }
  if (b.type === 'finding') {
    const auto = severityFor(cleanText(b.text + ' ' + b.cause), 'مشكلة');
    return `<section class="block block-finding sev-${b.severity === 'auto' ? 'auto' : b.severity}" id="block-${b.id}">${head('مؤشر (ملاحظة)', 'flag')}
      <div class="finding-grid">
        <label><span>مستوى الخطورة</span>${selectHtml(fid(b, 'severity'), f('severity'), [['auto', `تلقائي (${auto})`], ...SEVERITY_CHOICES.map(x => [x, x])], b.severity)}</label>
        <label><span>القسم</span>${selectHtml(fid(b, 'category'), f('category'), [['auto', 'تلقائي'], ...CATEGORY_NAMES.map(x => [x, x])], b.category)}</label>
        <label><span>الحالة</span>${selectHtml(fid(b, 'status'), f('status'), STATUS_OPTIONS, b.status)}</label>
        <label><span>المسؤول عن المعالجة</span><input id="${fid(b, 'owner')}" ${f('owner')} value="${h(b.owner)}" placeholder="الاسم أو الجهة"></label>
        <label><span>الموعد المستهدف</span><input type="date" id="${fid(b, 'due')}" ${f('due')} value="${h(b.due)}"></label></div>
      <label class="stack-label"><span>الملاحظة</span><textarea class="autosize block-area" id="${fid(b, 'text')}" ${f('text')} rows="2" placeholder="ما الذي رُصد؟">${h(b.text)}</textarea></label>
      <label class="stack-label"><span>السبب المحتمل <em>اختياري</em></span><textarea class="autosize block-area" id="${fid(b, 'cause')}" ${f('cause')} rows="1" placeholder="لماذا حدث؟">${h(b.cause)}</textarea></label>
      <label class="stack-label"><span>الإجراء المقترح <em>اختياري</em></span><textarea class="autosize block-area" id="${fid(b, 'action')}" ${f('action')} rows="1" placeholder="ما المطلوب فعله؟">${h(b.action)}</textarea></label></section>`;
  }
  if (b.type === 'image') {
    const p = currentProject();
    const src = state.assetUrls[b.assetId];
    return `<section class="block block-image" id="block-${b.id}">${head('صورة', 'image', `<button class="mini-btn" data-action="annotate" data-id="${b.id}">${icon('pen', 13)}تعليم وقص</button>`)}
      <div class="block-image-stage">${src ? `<img src="${h(src)}" alt="${h(b.caption || p?.imageAssets?.[b.assetId]?.fileName || 'صورة')}">` : `<div class="slide-image-unavailable"><span>جارٍ تحميل الصورة…</span></div>`}</div>
      <input class="caption-input" id="${fid(b, 'caption')}" ${f('caption')} value="${h(b.caption)}" placeholder="تعليق على الصورة (اختياري)"></section>`;
  }
  if (b.type === 'comment') {
    return `<section class="block block-comment" id="block-${b.id}">${head('تعليق', 'message', `<time class="block-time">${h(formatDate(b.createdAt))}</time>`)}
      <textarea class="autosize block-area" id="${fid(b, 'text')}" ${f('text')} rows="2" placeholder="اكتب تعليقك أو قرارك أو سؤالك…">${h(b.text)}</textarea>
      <input class="caption-input" id="${fid(b, 'author')}" ${f('author')} value="${h(b.author)}" placeholder="كاتب التعليق (اختياري)"></section>`;
  }
  if (b.type === 'table') {
    const rows = b.rows.map((row, r) => `<tr>${row.map((cell, c) => `<td><input id="cell-${b.id}-${r}-${c}" data-block="${b.id}" data-cell="${r},${c}" value="${h(cell)}" aria-label="الصف ${r + 1} العمود ${c + 1}"></td>`).join('')}</tr>`).join('');
    return `<section class="block block-table" id="block-${b.id}">${head('جدول', 'table')}
      <div class="table-scroll"><table class="edit-table"><tbody>${rows}</tbody></table></div>
      <div class="table-tools"><button class="text-button" data-action="table-add-row" data-id="${b.id}">${icon('plus', 13)}صف</button><button class="text-button" data-action="table-add-col" data-id="${b.id}">${icon('plus', 13)}عمود</button>
      <button class="text-button" data-action="table-del-row" data-id="${b.id}"${b.rows.length < 2 ? ' disabled' : ''}>حذف صف</button><button class="text-button" data-action="table-del-col" data-id="${b.id}"${(b.rows[0]?.length || 0) < 2 ? ' disabled' : ''}>حذف عمود</button>
      <small>تلميح: اجعل رأس الجدول «الملاحظة | السبب المحتمل | الإجراء المقترح» ليحلل النظام كل صف كحالة.</small></div></section>`;
  }
  return '';
}

function addBarHtml() {
  const btn = (type, ico, label) => `<button class="add-block" data-action="add-block" data-type="${type}">${icon(ico, 16)}<span>${label}</span></button>`;
  return `<div class="add-bar" role="group" aria-label="إضافة عنصر"><span class="add-bar-label">أضف إلى الشريحة</span>
    ${btn('text', 'type', 'نص')}${btn('finding', 'flag', 'مؤشر بمستوى خطورة')}
    <button class="add-block" data-action="add-image">${icon('imagePlus', 16)}<span>صورة</span></button>
    ${btn('comment', 'message', 'تعليق')}${btn('table', 'table', 'جدول')}</div>`;
}

function editorHtml(p) {
  const slide = p.slides[state.slideIdx] || p.slides[0];
  if (!slide) return '';
  const s = analysisOf(p).summary;
  const blocks = slide.blocks.map((b, i) => blockHtml(b, i, slide.blocks.length)).join('');
  return `<div class="editor-layout">
    <aside class="slide-strip" aria-label="شرائح المشروع"><div class="strip-head"><strong>الشرائح <span>${p.slides.length}</span></strong><button class="mini-primary" data-action="add-slide" title="شريحة جديدة" aria-label="شريحة جديدة">${icon('plus', 15)}</button></div>
      <div class="strip-list" id="slide-strip-list">${slideStripItems(p)}</div></aside>
    <section class="canvas" id="canvas" aria-label="الشريحة ${state.slideIdx + 1}">
      <div class="canvas-top"><span class="canvas-num">${pad2(state.slideIdx + 1)}</span>
        <input class="slide-title-input" id="slide-title" data-slide-title="${slide.id}" value="${h(slide.title)}" placeholder="عنوان الشريحة" aria-label="عنوان الشريحة">
        <div class="canvas-tools"><button data-action="move-slide" data-dir="-1" title="تقديم الشريحة" aria-label="تقديم الشريحة"${state.slideIdx === 0 ? ' disabled' : ''}>${icon('arrowUp', 15)}</button>
          <button data-action="move-slide" data-dir="1" title="تأخير الشريحة" aria-label="تأخير الشريحة"${state.slideIdx === p.slides.length - 1 ? ' disabled' : ''}>${icon('arrowDown', 15)}</button>
          <button data-action="dup-slide" title="نسخ الشريحة" aria-label="نسخ الشريحة">${icon('copy', 15)}</button>
          <button data-action="ask-delete-slide" class="danger" title="حذف الشريحة" aria-label="حذف الشريحة"${p.slides.length < 2 ? ' disabled' : ''}>${icon('trash', 15)}</button></div></div>
      <div class="live-strip" aria-live="polite"><span class="live-label">${icon('barChart', 14)}التحليل المباشر</span>
        <span class="risk-badge ${RISK_TONE[s.risk]}" id="live-risk">${s.risk}</span>
        <span><b id="live-cases">${s.cases.length}</b> حالة</span><span><b id="live-high">${s.bySeverity['مرتفع']}</b> مرتفعة</span><span><b id="live-open">${s.cases.length - s.byStatus.done}</b> مفتوحة</span>
        <span class="view-toggle" role="tablist" aria-label="طريقة التحرير"><button class="${state.editView === 'stage' ? 'on' : ''}" data-action="edit-view" data-view="stage" role="tab">${icon('monitor', 13)}لوحة الشريحة</button><button class="${state.editView === 'list' ? 'on' : ''}" data-action="edit-view" data-view="list" role="tab">${icon('layoutList', 13)}قائمة</button></span>
        <button class="text-button" data-action="mode-analyze">عرض التحليل الكامل ${icon('chevronLeft', 13)}</button></div>
      ${state.editView === 'stage' ? stageEditorBody(p, slide) : `<div class="blocks" id="blocks">${blocks || `<div class="blocks-empty">${icon('layers', 20)}<p>الشريحة فارغة. أضف نصًا أو مؤشرًا أو صورة أو تعليقًا من الأزرار أدناه، أو الصق صورة / اسحبها هنا.</p></div>`}</div>
      ${addBarHtml()}`}
      <details class="speaker-notes"${slide.notes ? ' open' : ''}><summary>${icon('mic', 14)}ملاحظات المتحدث <small>تُحلَّل مع الشريحة</small></summary>
        <textarea class="autosize block-area" id="slide-notes" data-slide-notes="${slide.id}" rows="2" placeholder="ما الذي ستقوله عند عرض هذه الشريحة؟">${h(slide.notes || '')}</textarea></details>
    </section></div>`;
}

/* ---------- العمليات على النموذج ---------- */
const curSlide = () => currentProject()?.slides[state.slideIdx];
function findBlock(id) { const slide = curSlide(); return slide ? { slide, index: slide.blocks.findIndex(b => b.id === id) } : null; }
function mutate(fn, { focus } = {}) {
  const p = currentProject(); if (!p) return;
  checkpoint();
  fn(p);
  touch(p);
  if (focus) state.focusId = focus;
  render();
}
function addBlock(type) {
  mutate(() => {
    const slide = curSlide();
    const b = newBlock(type, type === 'comment' ? { author: currentProject().owner || '' } : {});
    slide.blocks.push(b);
    state.sel = b.id;
    if (state.editView === 'stage' && (type === 'text' || type === 'comment')) { state.editingId = b.id; state.focusId = `inl-${b.id}`; }
    else state.focusId = type === 'table' ? `cell-${b.id}-0-0` : fid(b, 'text');
  });
}
function prepareImage(file) {
  return (async () => {
    if (!/^image\//.test(file.type)) throw new Error('اختر ملف صورة صالحًا.');
    if (file.size > 30 * 1024 * 1024) throw new Error('حجم الصورة أكبر من 30 ميغابايت.');
    if (/gif|svg/.test(file.type)) return file;
    try {
      const bmp = await createImageBitmap(file);
      const max = 1600, scale = Math.min(1, max / Math.max(bmp.width, bmp.height));
      if (scale === 1 && file.size < 600 * 1024) { bmp.close?.(); return file; }
      const canvas = document.createElement('canvas');
      canvas.width = Math.round(bmp.width * scale); canvas.height = Math.round(bmp.height * scale);
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(bmp, 0, 0, canvas.width, canvas.height); bmp.close?.();
      const blob = await new Promise(r => canvas.toBlob(r, 'image/jpeg', 0.86));
      return blob && blob.size < file.size ? blob : file;
    } catch { return file; }
  })();
}
async function addImageFiles(files) {
  const p = currentProject(), slide = curSlide();
  if (!p || !slide || state.mode !== 'edit') return;
  const images = Array.from(files).filter(f => /^image\//.test(f.type));
  if (!images.length) { state.notice = 'اختر ملف صورة صالحًا.'; render(); return; }
  let last = '';
  for (const file of images) {
    try {
      const blob = await prepareImage(file);
      const assetId = uid();
      await putAssetBlob(p.id, assetId, blob);
      state.assetUrls[assetId] = URL.createObjectURL(blob);
      p.imageAssets[assetId] = { id: assetId, fileName: file.name || 'صورة', contentType: blob.type || file.type, stored: true };
      const b = newBlock('image', { assetId });
      slide.blocks.push(b); last = b.id;
    } catch (reason) { state.error = reason instanceof Error ? reason.message : 'تعذّر حفظ الصورة.'; }
  }
  touch(p);
  if (last) { state.sel = last; if (state.editView !== 'stage') state.focusId = fid({ id: last }, 'caption'); }
  render();
}
function cloneBlock(b) { return { ...JSON.parse(JSON.stringify(b)), id: uid() }; }
function cloneSlide(slide) { return { ...JSON.parse(JSON.stringify(slide)), id: uid(), title: slide.title ? `${slide.title} (نسخة)` : '', blocks: slide.blocks.map(cloneBlock) }; }

/** أي حقل يحمل data-block + data-field يُحدّث النموذج مباشرة دون إعادة رسم. */
function applyField(t) {
  const p = currentProject(); if (!p) return false;
  if (t.dataset.slideTitle) { const s = p.slides.find(x => x.id === t.dataset.slideTitle); if (s) { s.title = t.value; touch(p); } return true; }
  if (t.dataset.slideNotes) { const s = p.slides.find(x => x.id === t.dataset.slideNotes); if (s) { s.notes = t.value; touch(p); } return true; }
  if (t.dataset.block) {
    const slide = curSlide(); const b = slide?.blocks.find(x => x.id === t.dataset.block);
    if (!b) return false;
    if (t.dataset.cell) { const [r, c] = t.dataset.cell.split(',').map(Number); b.rows[r][c] = t.value; }
    else if (t.dataset.field) b[t.dataset.field] = t.value;
    touch(p); return true;
  }
  return false;
}
function autosizeAll() {
  document.querySelectorAll('textarea.autosize').forEach(el => { el.style.height = 'auto'; el.style.height = `${el.scrollHeight + 2}px`; });
}

const editorActions = {
  'select-slide': (t) => { state.slideIdx = Number(t.dataset.index); state.sel = null; state.editingId = null; render(); document.getElementById('canvas')?.scrollIntoView({ block: 'nearest' }); },
  'add-slide': () => mutate(p => { p.slides.splice(state.slideIdx + 1, 0, newSlide(`الشريحة ${p.slides.length + 1}`)); state.slideIdx += 1; state.sel = null; state.focusId = 'slide-title'; }),
  'dup-slide': () => mutate(p => { p.slides.splice(state.slideIdx + 1, 0, cloneSlide(p.slides[state.slideIdx])); state.slideIdx += 1; }),
  'move-slide': (t) => mutate(p => {
    const to = state.slideIdx + Number(t.dataset.dir);
    if (to < 0 || to >= p.slides.length) return;
    const [slide] = p.slides.splice(state.slideIdx, 1); p.slides.splice(to, 0, slide); state.slideIdx = to;
  }),
  'ask-delete-slide': () => {
    const p = currentProject(); if (!p || p.slides.length < 2) return;
    const slide = p.slides[state.slideIdx];
    if (!slide.blocks.length && !slide.title.trim()) { actions['confirm-delete-slide'](); return; }
    state.deleteSlide = slide; render();
  },
  'confirm-delete-slide': () => {
    state.deleteSlide = null;
    mutate(p => { p.slides.splice(state.slideIdx, 1); state.slideIdx = Math.min(state.slideIdx, p.slides.length - 1); });
    state.notice = 'تم حذف الشريحة.'; render();
  },
  'add-block': (t) => addBlock(t.dataset.type),
  'add-image': () => $imageInput.click(),
  'move-block': (t) => mutate(() => {
    const { slide, index } = findBlock(t.dataset.id); const to = index + Number(t.dataset.dir);
    if (index < 0 || to < 0 || to >= slide.blocks.length) return;
    const [b] = slide.blocks.splice(index, 1); slide.blocks.splice(to, 0, b);
  }),
  'dup-block': (t) => mutate(() => { const { slide, index } = findBlock(t.dataset.id); if (index < 0) return; const c = cloneBlock(slide.blocks[index]); if (c.box) { c.box.x = Math.min(100 - c.box.w, c.box.x + 2.5); c.box.y = Math.min(96, c.box.y + 3); c.box.auto = false; } slide.blocks.splice(index + 1, 0, c); state.sel = c.id; }),
  'del-block': (t) => {
    const { slide, index } = findBlock(t.dataset.id); if (index < 0) return;
    const block = slide.blocks[index];
    clearTimeout(undoTimer);
    state.undo = { slideId: slide.id, block, index };
    undoTimer = window.setTimeout(() => { state.undo = null; render(); }, 9000);
    mutate(() => { slide.blocks.splice(index, 1); });
  },
  'undo': () => {
    const u = state.undo; if (!u) return;
    clearTimeout(undoTimer); state.undo = null;
    mutate(p => { const slide = p.slides.find(s => s.id === u.slideId); if (slide) slide.blocks.splice(Math.min(u.index, slide.blocks.length), 0, u.block); });
  },
  'table-add-row': (t) => mutate(() => { const b = curSlide().blocks.find(x => x.id === t.dataset.id); b.rows.push(b.rows[0].map(() => '')); }),
  'table-add-col': (t) => mutate(() => { const b = curSlide().blocks.find(x => x.id === t.dataset.id); b.rows.forEach(r => r.push('')); }),
  'table-del-row': (t) => mutate(() => { const b = curSlide().blocks.find(x => x.id === t.dataset.id); if (b.rows.length > 1) b.rows.pop(); }),
  'table-del-col': (t) => mutate(() => { const b = curSlide().blocks.find(x => x.id === t.dataset.id); if (b.rows[0].length > 1) b.rows.forEach(r => r.pop()); }),
};
let undoTimer = 0;

/* ---------- السحب والإفلات: ترتيب الشرائح والعناصر، ونقل عنصر إلى شريحة أخرى ---------- */
function clearDropMarks() { document.querySelectorAll('.drop-before,.drop-after,.drop-into,.dragging').forEach(el => el.classList.remove('drop-before', 'drop-after', 'drop-into', 'dragging')); }
function dropTarget(event) {
  const d = state.drag; if (!d) return null;
  const el = event.target.closest?.(d.kind === 'slide' ? '.strip-item' : '.block, .strip-item');
  if (!el) return null;
  const rect = el.getBoundingClientRect();
  const horizontal = el.classList.contains('strip-item') && rect.width > rect.height * 1.6 === false && window.matchMedia('(max-width: 980px)').matches;
  const after = horizontal ? event.clientX < rect.left + rect.width / 2 : event.clientY > rect.top + rect.height / 2;
  if (el.classList.contains('strip-item')) return { el, kind: 'slide', index: Number(el.dataset.dragId), after, into: d.kind === 'block' };
  return { el, kind: 'block', id: el.id.replace('block-', ''), after };
}
function dropSlide(from, to, after) {
  const p = currentProject(); if (!p || from === to) return;
  const currentId = p.slides[state.slideIdx]?.id;
  const [slide] = p.slides.splice(from, 1);
  let target = to > from ? to - 1 : to; if (after) target += 1;
  p.slides.splice(target, 0, slide);
  state.slideIdx = Math.max(0, p.slides.findIndex(s => s.id === currentId));
  touch(p); render();
}
function dropBlock(id, t) {
  const p = currentProject(), slide = curSlide(); if (!p || !slide) return;
  const from = slide.blocks.findIndex(b => b.id === id); if (from < 0) return;
  if (t.kind === 'slide') {
    const dest = p.slides[t.index]; if (!dest || dest === slide) return;
    const [b] = slide.blocks.splice(from, 1); dest.blocks.push(b);
    touch(p); state.notice = `نُقل العنصر إلى الشريحة ${t.index + 1}.`; render(); return;
  }
  if (t.id === id) return;
  const [b] = slide.blocks.splice(from, 1);
  let to = slide.blocks.findIndex(x => x.id === t.id); if (t.after) to += 1;
  slide.blocks.splice(to, 0, b);
  touch(p); render();
}
document.addEventListener('dragstart', (e) => {
  const src = e.target.closest?.('[data-drag-kind]'); if (!src) return;
  state.drag = { kind: src.dataset.dragKind, id: src.dataset.dragId };
  e.dataTransfer.effectAllowed = 'move';
  try { e.dataTransfer.setData('text/plain', `waraq:${state.drag.kind}:${state.drag.id}`); } catch { /* بعض المتصفحات */ }
  const shown = state.drag.kind === 'block' ? document.getElementById(`block-${state.drag.id}`) : src;
  if (shown) { try { e.dataTransfer.setDragImage(shown, 10, 10); } catch { /* اختياري */ } shown.classList.add('dragging'); }
});
document.addEventListener('dragover', (e) => {
  if (!state.drag) return;
  const t = dropTarget(e); clearDropMarks();
  if (!t) return;
  e.preventDefault(); e.dataTransfer.dropEffect = 'move';
  t.el.classList.add(t.into ? 'drop-into' : t.after ? 'drop-after' : 'drop-before');
});
document.addEventListener('drop', (e) => {
  const d = state.drag; if (!d) return;
  const t = dropTarget(e); clearDropMarks(); state.drag = null;
  if (!t) return;
  e.preventDefault(); e.stopPropagation();
  if (d.kind === 'slide') dropSlide(Number(d.id), t.index, t.after); else dropBlock(d.id, t);
}, true);
document.addEventListener('dragend', () => { clearDropMarks(); state.drag = null; });

/* ================= القائمة اليمنى: المشاريع والتحليلات المحفوظة ================= */
function relativeDay(value) {
  const days = Math.floor((Date.now() - new Date(value).getTime()) / 86400000);
  if (days <= 0) return 'اليوم';
  if (days === 1) return 'أمس';
  return formatDate(value);
}
function filteredProjects() {
  const q = state.projectSearch.trim().toLocaleLowerCase('ar');
  return state.projects
    .filter(p => !q || `${p.name} ${p.owner} ${p.location}`.toLocaleLowerCase('ar').includes(q))
    .slice().sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
}
const onProject = (p) => state.view === 'project' && p.id === state.projectId;
/** قسم «التحليلات»: تحليل كل مشروع محفوظ (المستوى والحالات والتقدم وآخر تحليل). */
function analysisListHtml() {
  const list = filteredProjects();
  if (!list.length) return `<p class="side-empty">${state.projects.length ? 'لا تحليلات بهذا البحث.' : 'تظهر هنا تحليلات مشاريعك بعد إنشائها.'}</p>`;
  return list.map(p => {
    const s = analysisOf(p).summary, last = (p.history || []).at(-1)?.at || p.updatedAt;
    return `<button class="an-item${onProject(p) && state.mode === 'analyze' ? ' an-active' : ''}" data-action="open-project" data-id="${h(p.id)}" data-mode="analyze" aria-label="تحليل ${h(p.name)}">
      <span class="an-top"><span class="proj-name">${h(p.name)}</span><span class="risk-badge tiny ${RISK_TONE[s.risk]}">${s.risk}</span></span>
      <span class="progress-stack mini" aria-hidden="true"><i class="ps-done" style="width:${s.progress}%"></i></span>
      <span class="proj-stats"><span>${s.cases.length} حالة</span>${s.bySeverity['مرتفع'] ? `<span class="hot-count">${s.bySeverity['مرتفع']} مرتفعة</span>` : ''}<span>${s.progress}% مُنجز</span><span class="muted-s">${h(relativeDay(last))}</span></span></button>`;
  }).join('');
}
function projectListHtml() {
  const list = filteredProjects();
  if (!list.length) return `<p class="side-empty">${state.projects.length ? 'لا مشاريع بهذا البحث.' : 'لا توجد مشاريع محفوظة بعد.'}</p>`;
  return list.map(p => {
    const on = onProject(p);
    return `<div class="proj-item${on ? ' proj-active' : ''}" data-project="${h(p.id)}">
      <button class="proj-main" data-action="open-project" data-id="${h(p.id)}" data-mode="edit" aria-label="فتح ${h(p.name)}"${on ? ' aria-current="true"' : ''}>
        <span class="proj-name">${h(p.name)}</span>
        <span class="proj-meta">${p.slides.length} شريحة · ${h(relativeDay(p.updatedAt))}${p.owner ? ` · ${h(p.owner)}` : ''}</span></button>
      <div class="proj-actions"><button data-action="open-project" data-id="${h(p.id)}" data-mode="edit" title="تحرير المشروع" aria-label="تحرير ${h(p.name)}">${icon('pencil', 14)}</button>
        <button data-action="dup-project" data-id="${h(p.id)}" title="نسخ المشروع" aria-label="نسخ ${h(p.name)}">${icon('copy', 14)}</button>
        <button data-action="ask-delete-project" data-id="${h(p.id)}" class="danger" title="حذف المشروع" aria-label="حذف ${h(p.name)}">${icon('trash', 14)}</button></div></div>`;
  }).join('');
}
function sidebarHtml() {
  const all = state.projects.map(p => analysisOf(p).summary);
  const totalCases = all.reduce((n, s) => n + s.cases.length, 0);
  const openHigh = all.reduce((n, s) => n + s.openHigh, 0);
  return `<aside class="sidebar${state.sidebarOpen ? ' sidebar-open' : ''}" id="sidebar" aria-label="التحليلات والمشاريع">
    <div class="side-head"><h2>وَرَق</h2><button class="side-fold desktop-only" data-action="toggle-side-collapse" title="إخفاء القائمة لتوسيع مساحة العمل">إخفاء ${icon('chevronRight', 13)}</button><button class="side-close mobile-only-label" data-action="toggle-sidebar" aria-label="إغلاق القائمة">${icon('x', 18)}</button></div>
    <div class="side-overview" aria-label="نظرة عامة على كل المشاريع"><div><strong>${state.projects.length}</strong><span>مشروع</span></div><div><strong>${totalCases}</strong><span>حالة محللة</span></div><div class="${openHigh ? 'hot' : ''}"><strong>${openHigh}</strong><span>مرتفعة مفتوحة</span></div></div>
    <button class="board-link${state.view === 'board' ? ' board-on' : ''}" data-action="open-board">${icon('calendar', 16)}لوحة المتابعة والمواعيد<span class="board-count">${boardOverdueCount()}</span></button>
    <label class="search-field side-search">${icon('search')}<input id="project-search" value="${h(state.projectSearch)}" placeholder="ابحث في التحليلات والمشاريع..." aria-label="ابحث في التحليلات والمشاريع"></label>
    <div class="side-title">${icon('barChart', 13)}التحليلات المحفوظة</div>
    <div class="analysis-list" id="analysis-list">${analysisListHtml()}</div>
    <div class="side-title with-action"><span>${icon('folder', 13)}مشاريعي</span><button class="side-mini" data-action="new-project">${icon('plus', 13)}جديد</button></div>
    <div class="project-list" id="project-list">${projectListHtml()}</div>
    <button class="primary-button side-new" data-action="new-project">${icon('plus', 17)}مشروع جديد</button>
    <button class="side-import" data-action="import">${icon('upload', 15)}استيراد (PowerPoint أو نسخة احتياطية)</button>
    <label class="side-option" title="الافتراضي: الصور الأساسية فقط. فعّله لاستيراد كل صور ملف PowerPoint."><input type="checkbox" id="opt-all-images"${state.allImages ? ' checked' : ''}><span>استيراد كل الصور من PowerPoint</span></label>
    <div class="side-foot">${icon('shield', 13)}<span>كل شيء محفوظ في هذا المتصفح فقط</span></div></aside>`;
}

/* ================= لوحة المتابعة (كل المشاريع) ================= */
const todayStart = () => { const d = new Date(); d.setHours(0, 0, 0, 0); return d.getTime(); };
const daysUntil = (due) => due ? Math.round((new Date(`${due}T00:00:00`).getTime() - todayStart()) / 86400000) : null;
function dueLabel(d) { if (d === null) return 'بلا موعد'; if (d < 0) return `متأخرة ${-d} يوم`; if (d === 0) return 'اليوم'; return d === 1 ? 'غدًا' : `بعد ${d} يوم`; }
function boardRows() {
  const rows = [];
  for (const p of state.projects) for (const c of analysisOf(p).summary.cases) rows.push({ p, c, days: daysUntil(c.due), owner: c.owner || 'غير مُسنَد' });
  return rows;
}
function boardOverdueCount() { return boardRows().filter(r => r.c.status !== 'done' && r.days !== null && r.days < 0).length; }
function boardHtml() {
  const all = boardRows(), open = all.filter(r => r.c.status !== 'done');
  const owners = Array.from(new Set(all.map(r => r.owner))).sort((a, b) => a.localeCompare(b, 'ar'));
  const scopes = { all: 'كل المفتوحة', overdue: 'المتأخرة', week: 'خلال ٧ أيام', nodue: 'بلا موعد' };
  const inScope = (r) => state.boardScope === 'all' || (state.boardScope === 'overdue' && r.days !== null && r.days < 0) || (state.boardScope === 'week' && r.days !== null && r.days >= 0 && r.days <= 7) || (state.boardScope === 'nodue' && r.days === null);
  const list = open.filter(r => inScope(r) && (state.boardOwner === 'الكل' || r.owner === state.boardOwner) && (state.boardProject === 'الكل' || r.p.id === state.boardProject))
    .sort((a, b) => (a.days ?? 9999) - (b.days ?? 9999) || SEVERITY_RANK[a.c.severity] - SEVERITY_RANK[b.c.severity]);
  const overdue = open.filter(r => r.days !== null && r.days < 0).length, week = open.filter(r => r.days !== null && r.days >= 0 && r.days <= 7).length;
  const stat = (n, label, tone = '') => `<div class="board-stat ${tone}"><strong>${n}</strong><span>${label}</span></div>`;
  const load = owners.map(o => { const mine = all.filter(r => r.owner === o), op = mine.filter(r => r.c.status !== 'done'); return { o, open: op.length, high: op.filter(r => r.c.severity === 'مرتفع').length, late: op.filter(r => r.days !== null && r.days < 0).length, done: mine.length - op.length }; })
    .sort((a, b) => b.open - a.open);
  const opt = (v, l, cur) => `<option value="${h(v)}"${v === cur ? ' selected' : ''}>${h(l)}</option>`;
  return `<section class="project-bar animate-in"><div class="pb-top"><div class="pb-title"><div class="eyebrow"><span class="eyebrow-rule"></span>كل المشاريع</div><h1>لوحة المتابعة والمواعيد</h1><p class="pb-desc">الحالات المفتوحة في كل مشاريعك: ما تأخر موعده، وما يستحق خلال أسبوع، وما على كل مسؤول.</p></div></div></section>
  <section class="board-stats animate-in delay-1">${stat(open.length, 'حالة مفتوحة')}${stat(overdue, 'متأخرة', overdue ? 'bad' : '')}${stat(week, 'خلال ٧ أيام', week ? 'warn' : '')}${stat(open.filter(r => r.c.severity === 'مرتفع').length, 'مرتفعة الخطورة', 'bad')}${stat(open.filter(r => r.days === null).length, 'بلا موعد')}</section>
  <section class="panel board-panel animate-in delay-2"><div class="content-header"><div><span class="section-kicker">الحالات المفتوحة</span><h2>حسب الموعد</h2></div><div class="count-badge">${list.length}</div></div>
    <div class="filters">
      <label class="filter-select"><span>النطاق</span><select id="board-scope">${Object.entries(scopes).map(([k, v]) => opt(k, v, state.boardScope)).join('')}</select>${icon('chevronDown', 13)}</label>
      <label class="filter-select"><span>المسؤول</span><select id="board-owner">${['الكل', ...owners].map(o => opt(o, o, state.boardOwner)).join('')}</select>${icon('chevronDown', 13)}</label>
      <label class="filter-select"><span>المشروع</span><select id="board-project">${opt('الكل', 'الكل', state.boardProject)}${state.projects.map(p => opt(p.id, p.name, state.boardProject)).join('')}</select>${icon('chevronDown', 13)}</label></div>
    <div class="table-scroll"><table class="board-table"><thead><tr><th>الموعد</th><th>الحالة</th><th>الخطورة</th><th>المشروع · الشريحة</th><th>المسؤول</th><th>الحالة المطلوبة</th><th></th></tr></thead><tbody>
    ${list.map(r => `<tr class="${r.days !== null && r.days < 0 ? 'late' : ''}"><td><span class="due-chip ${r.days === null ? 'none' : r.days < 0 ? 'late' : r.days <= 7 ? 'soon' : ''}">${h(dueLabel(r.days))}</span>${r.c.due ? `<small>${h(r.c.due)}</small>` : ''}</td>
      <td><select class="board-status" data-project="${h(r.p.id)}" data-case="${h(r.c.id)}" aria-label="حالة المتابعة">${Object.entries(STATUS_LABELS).map(([k, v]) => `<option value="${k}"${r.c.status === k ? ' selected' : ''}>${v}</option>`).join('')}</select></td>
      <td><span class="severity-tag ${severityTone[r.c.severity]}">${r.c.severity}</span></td><td><b>${h(r.p.name)}</b><small>شريحة ${r.c.slideNumber} · ${h(r.c.section)}</small></td><td>${h(r.owner)}</td><td class="board-title">${h(r.c.title.slice(0, 110))}${r.c.title.length > 110 ? '…' : ''}</td>
      <td><button class="text-button" data-action="board-open" data-project="${h(r.p.id)}" data-slide="${r.c.slideNumber}" data-block="${h(r.c.blockId || '')}">${icon('pencil', 13)}فتح</button></td></tr>`).join('') || '<tr><td colspan="7" class="side-empty">لا حالات في هذا النطاق.</td></tr>'}</tbody></table></div></section>
  <section class="panel board-panel animate-in delay-3"><div class="content-header"><div><span class="section-kicker">الأحمال</span><h2>ما على كل مسؤول</h2></div></div>
    <div class="table-scroll"><table class="board-table"><thead><tr><th>المسؤول</th><th>مفتوحة</th><th>مرتفعة</th><th>متأخرة</th><th>مغلقة</th></tr></thead><tbody>
    ${load.map(x => `<tr><td><button class="text-button" data-action="board-owner" data-owner="${h(x.o)}">${h(x.o)}</button></td><td>${x.open}</td><td class="${x.high ? 'hot-count' : ''}">${x.high}</td><td class="${x.late ? 'hot-count' : ''}">${x.late}</td><td>${x.done}</td></tr>`).join('') || '<tr><td colspan="5" class="side-empty">لا بيانات بعد.</td></tr>'}</tbody></table></div></section>`;
}

/* ================= شريط المشروع ================= */
function projectBarHtml(p) {
  const edit = state.mode === 'edit';
  return `<section class="project-bar animate-in${edit ? ' compact' : ''}">
    <div class="pb-top"><div class="pb-title"><div class="eyebrow"><span class="eyebrow-rule"></span>${edit ? 'وضع التحرير' : 'وضع التحليل'}</div><h1>${h(p.name)}</h1>
      <p class="pb-meta">${p.owner ? `${icon('user', 13)}${h(p.owner)}` : ''}${p.location ? `${icon('folder', 13)}${h(p.location)}` : ''}<span>${icon('layers', 13)}${p.slides.length} شريحة</span><span>${icon('clock', 13)}آخر تعديل ${h(relativeDay(p.updatedAt))}</span></p>
      ${p.description ? `<p class="pb-desc">${h(p.description)}</p>` : ''}</div>
      <button class="text-button" data-action="edit-info">${icon('pencil', 14)}معلومات المشروع</button></div>
    <div class="pb-actions">
      <div class="mode-switch" role="tablist"><button role="tab" class="${edit ? 'on' : ''}" data-action="mode-edit">${icon('pencil', 15)}تحرير</button><button role="tab" class="${edit ? '' : 'on'}" data-action="mode-analyze">${icon('barChart', 15)}التحليل</button></div>
      <button class="top-action light" data-action="mode-present">${icon('play', 14)}عرض تقديمي</button>
      <span class="save-indicator" id="save-indicator"></span><span class="grow"></span>
      ${edit ? `<button class="primary-button" data-action="save-analyze">${icon('save', 16)}حفظ وتحليل</button>` : ''}
      <button class="top-action light" data-action="export-pptx" title="تصدير الشرائح كملف PowerPoint">${icon('presentation', 14)}PowerPoint</button>
      <button class="top-action light" data-action="export-html">${icon('download', 14)}تقرير</button>
      <button class="top-action light" data-action="export-csv">${icon('table', 14)}CSV</button>
      <button class="top-action light" data-action="export-backup" title="نسخة كاملة مع الصور تُستورد لاحقًا">${icon('fileJson', 14)}نسخة احتياطية</button></div></section>`;
}

/* ================= واجهة التحليل ================= */
function commentsTabHtml(a) {
  if (!a.notes.length) return `<div class="notes-empty"><div class="note-empty-mark">${icon('message', 18)}</div><p>لا توجد تعليقات.<br>أضف «تعليقًا» إلى أي شريحة من المحرّر.</p></div>`;
  return `<div class="notes-list wide">${a.notes.map(n => `<article class="note-card"><div class="note-card-top"><span>شريحة ${n.slideNumber}${n.author ? ` · ${h(n.author)}` : ''}</span>
    <button data-action="goto-comment" data-id="${h(n.id)}" data-slide="${n.slideNumber}" aria-label="تحرير التعليق">${icon('pencil', 14)}</button></div><p>${h(n.text)}</p><time>${h(formatDate(n.createdAt))}</time></article>`).join('')}</div>`;
}
function historyTabHtml(p) {
  const list = (p.history || []).slice().reverse();
  if (!list.length) return `<div class="notes-empty"><div class="note-empty-mark">${icon('clock', 18)}</div><p>يُسجَّل تحليل جديد في كل مرة تفتح فيها وضع التحليل بعد تعديل المشروع.</p></div>`;
  return `<div class="history-list">${list.map((e, i) => {
    const prev = list[i + 1];
    const delta = prev ? e.total - prev.total : 0;
    return `<div class="history-row"><span class="history-date">${h(new Intl.DateTimeFormat('ar', { dateStyle: 'medium', timeStyle: 'short' }).format(new Date(e.at)))}</span>
      <span class="risk-badge tiny ${RISK_TONE[e.risk]}">${e.risk}</span><span>${e.total} حالة${prev && delta ? ` (${delta > 0 ? '+' : ''}${delta})` : ''}</span><span>${e.high} مرتفعة</span><span>${e.progress}% مُنجز</span><span class="muted-s">${e.slides} شريحة</span></div>`;
  }).join('')}</div>`;
}
function workspaceHtml(p, a, s) {
  const imgCount = galleryItems(a).length;
  const tab = state.tab;
  const count = tab === 'cases' ? `${casesView(a, s).length} حالة` : tab === 'slides' ? `${a.slides.length} شريحة` : tab === 'images' ? `${imgCount} صورة` : tab === 'comments' ? `${a.notes.length} تعليق` : `${(p.history || []).length} تحليل`;
  const body = tab === 'cases' ? casesTabHtml(a, s) : tab === 'slides' ? slidesTabHtml(a, s) : tab === 'images' ? imagesTabHtml(a) : tab === 'comments' ? commentsTabHtml(a) : historyTabHtml(p);
  const tabBtn = (id, label, n) => `<button class="${tab === id ? 'tab-active' : ''}" data-action="tab" data-value="${id}" role="tab">${label} <span>${n}</span></button>`;
  return `<section class="panel findings-panel animate-in delay-3" id="workspace">
    <div class="content-header"><div><span class="section-kicker">نتائج التحليل</span><h2>سجلّ المراجعة</h2></div><div class="count-badge" id="count-badge">${count}</div></div>
    <div class="tabs-row" role="tablist">${tabBtn('cases', 'الحالات', s.cases.length)}${tabBtn('slides', 'الشرائح', a.slides.length)}${tabBtn('images', 'الصور', imgCount)}${tabBtn('comments', 'التعليقات', a.notes.length)}${tabBtn('history', 'سجل التحليلات', (p.history || []).length)}</div>
    ${body}<div class="panel-footer"><span>${icon('shield', 14)} المشروع محفوظ على هذا الجهاز فقط</span><button data-action="edit-info">تعديل معلومات المشروع ${icon('pencil', 14)}</button></div></section>`;
}

/* ================= العرض التقديمي ================= */
function presentBlock(b) {
  if (b.type === 'text') {
    if (b.style === 'h') return `<h2>${h(b.text)}</h2>`;
    if (b.style === 'ul') return `<ul>${b.text.split('\n').filter(l => l.trim()).map(l => `<li>${h(l)}</li>`).join('')}</ul>`;
    return b.text.trim() ? `<p>${h(b.text).replace(/\n/g, '<br>')}</p>` : '';
  }
  if (b.type === 'finding') {
    const sev = SEVERITY_CHOICES.includes(b.severity) ? b.severity : severityFor(cleanText(`${b.text} ${b.cause}`), 'مشكلة');
    if (!`${b.text}${b.cause}${b.action}`.trim()) return '';
    return `<div class="pres-finding ${severityTone[sev]}"><div class="pres-find-head"><span class="severity-tag ${severityTone[sev]}">${sev}</span>${b.owner ? `<span>${icon('user', 13)}${h(b.owner)}</span>` : ''}${b.due ? `<span>${icon('calendar', 13)}${h(b.due)}</span>` : ''}<span>${h(STATUS_LABELS[b.status] || '')}</span></div>
      <p><strong>${h(b.text)}</strong></p>${b.cause ? `<p><em>السبب المحتمل:</em> ${h(b.cause)}</p>` : ''}${b.action ? `<p class="pres-action"><em>الإجراء المقترح:</em> ${h(b.action)}</p>` : ''}</div>`;
  }
  if (b.type === 'image') { const src = state.assetUrls[b.assetId]; return src ? `<figure><img src="${h(src)}" alt="${h(b.caption || 'صورة')}">${b.caption ? `<figcaption>${h(b.caption)}</figcaption>` : ''}</figure>` : ''; }
  if (b.type === 'table') return `<div class="table-scroll"><table class="pres-table"><tbody>${b.rows.map((r, i) => `<tr>${r.map(c => i === 0 ? `<th>${h(c)}</th>` : `<td>${h(c)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
  return '';
}
function presentHtml(p) {
  ensureAllLayouts(p);
  const i = Math.min(state.present, p.slides.length - 1), slide = p.slides[i];
  return `<div class="present" role="dialog" aria-modal="true" aria-label="عرض تقديمي"><div class="present-top"><span>${h(p.name)}</span><span class="present-count">${i + 1} / ${p.slides.length}</span><button class="icon-button dark" data-action="present-close" aria-label="إغلاق العرض">${icon('x', 19)}</button></div>
    <div class="present-stagewrap">${stageHtml(slide, false)}</div>
    <div class="present-nav"><button data-action="present-prev" aria-label="السابقة"${i === 0 ? ' disabled' : ''}>${icon('chevronRight', 20)}السابقة</button><button data-action="present-next" aria-label="التالية"${i === p.slides.length - 1 ? ' disabled' : ''}>التالية${icon('chevronLeft', 20)}</button></div></div>`;
}

/* ================= النوافذ ================= */
function modalsHtml(p, a, s) {
  let out = '';
  if (a && (state.selectedFinding || state.selectedSlide)) {
    const f = state.selectedFinding, sl = state.selectedSlide;
    const number = f?.slideNumber || sl?.number || 0;
    const slide = a.slides.find(x => x.number === number);
    const c = f && s.cases.find(x => x.id === (f.caseId || f.id));
    out += `<div class="modal-backdrop" data-backdrop="close-quote"><section class="quote-modal" role="dialog" aria-modal="true" aria-labelledby="quote-title">
      <div class="modal-top"><div><span class="section-kicker">من المشروع</span><h2 id="quote-title">نص الشريحة</h2></div><button class="icon-button" data-action="close-quote" aria-label="إغلاق">${icon('x', 19)}</button></div>
      <div class="quote-source"><span class="slide-number">${pad2(number)}</span><div><strong>${h(slide?.title)}</strong><small>${h(a.fileName)}</small></div></div>
      ${c ? `<div class="finding-detail"><span class="severity-tag ${severityTone[c.severity]}">${c.severity}</span><p>${h(c.title)}</p></div>` : ''}
      <blockquote class="quote-full">${h(f?.evidence || sl?.text || 'لا يوجد نص في هذه الشريحة.')}</blockquote>
      ${slide ? galleryHtml(slide) : ''}
      <div class="full-slide-text">${slide ? h(slide.text.split('\n').slice(1).join('\n')) : ''}</div>
      ${slide?.notes ? `<div class="full-slide-text"><strong>ملاحظات المتحدث</strong>\n${h(slide.notes)}</div>` : ''}
      <div class="modal-actions"><button class="secondary-button" data-action="edit-from-quote">${icon('pencil')}تحرير هذه الشريحة</button><button class="text-button" data-action="close-quote">إغلاق</button></div></section></div>`;
  }
  const d = state.projectForm;
  if (d) {
    const isNew = d.mode === 'new';
    out += `<div class="modal-backdrop" data-backdrop="close-form"><section class="note-modal" role="dialog" aria-modal="true" aria-labelledby="pf-title">
      <div class="modal-top"><div><span class="section-kicker">${isNew ? 'بداية جديدة' : 'بيانات المشروع'}</span><h2 id="pf-title">${isNew ? 'مشروع جديد' : 'معلومات المشروع'}</h2></div><button class="icon-button" data-action="close-form" aria-label="إغلاق">${icon('x', 19)}</button></div>
      <label class="field-label" for="pf-name">اسم المشروع</label><input class="text-input" id="pf-name" value="${h(d.name)}" placeholder="مثال: مراجعة موقع الرياض" maxlength="120">
      <div class="two-cols"><div><label class="field-label" for="pf-owner">المسؤول <span>اختياري</span></label><input class="text-input" id="pf-owner" value="${h(d.owner)}" placeholder="اسم المسؤول"></div>
      <div><label class="field-label" for="pf-location">الموقع <span>اختياري</span></label><input class="text-input" id="pf-location" value="${h(d.location)}" placeholder="المدينة أو المبنى"></div></div>
      <label class="field-label" for="pf-description">وصف مختصر <span>اختياري</span></label><textarea id="pf-description" placeholder="نطاق المشروع وهدفه…">${h(d.description)}</textarea>
      ${isNew ? `<label class="field-label" for="pf-template">البداية</label><select class="text-input" id="pf-template"><option value="blank"${d.template === 'blank' ? ' selected' : ''}>مشروع فارغ (شريحة واحدة)</option><option value="review"${d.template === 'review' ? ' selected' : ''}>قالب مراجعة ميدانية (٤ شرائح جاهزة)</option></select>` : ''}
      <div class="modal-actions"><button class="primary-button" id="save-form" data-action="save-form"${d.name.trim() ? '' : ' disabled'}>${icon('check')}${isNew ? 'إنشاء المشروع' : 'حفظ'}</button><button class="text-button" data-action="close-form">إلغاء</button></div></section></div>`;
  }
  if (a && state.lightbox !== null) {
    const items = galleryItems(a), item = items[state.lightbox];
    if (item) {
      const src = imageSrc(a, item.id);
      out += `<div class="modal-backdrop lightbox" data-backdrop="close-lightbox"><section class="lightbox-card" role="dialog" aria-modal="true" aria-label="معاينة الصورة">
        <div class="lightbox-bar"><span>شريحة ${item.slide.number} · ${h(item.slide.title.slice(0, 60))} <small>(${state.lightbox + 1} / ${items.length})</small></span>
          <div class="row-2"><a class="top-action dark" href="${h(src)}" download="${h(item.asset.fileName)}">${icon('download', 15)}تنزيل</a><button class="icon-button" data-action="close-lightbox" aria-label="إغلاق">${icon('x', 19)}</button></div></div>
        <div class="lightbox-stage"><button class="lb-nav" data-action="lb-next" aria-label="التالية">${icon('chevronRight', 22)}</button><img src="${h(src)}" alt="صورة الشريحة ${item.slide.number}"><button class="lb-nav" data-action="lb-prev" aria-label="السابقة">${icon('chevronLeft', 22)}</button></div></section></div>`;
    }
  }
  if (state.deleteSlide) out += `<div class="modal-backdrop" data-backdrop="cancel-delete"><section class="confirm-modal" role="dialog" aria-modal="true"><div class="confirm-icon">${icon('trash', 19)}</div><h2>حذف هذه الشريحة؟</h2><p>ستُحذف الشريحة بكل ما فيها من نصوص ومؤشرات وصور وتعليقات.</p><div class="modal-actions"><button class="danger-button" data-action="confirm-delete-slide">حذف الشريحة</button><button class="text-button" data-action="cancel-delete">إبقاء</button></div></section></div>`;
  if (state.deleteProject) out += `<div class="modal-backdrop" data-backdrop="cancel-delete"><section class="confirm-modal" role="dialog" aria-modal="true"><div class="confirm-icon">${icon('trash', 19)}</div><h2>حذف «${h(state.deleteProject.name)}»؟</h2><p>سيُحذف المشروع وتحليله وصوره من هذا المتصفح. يمكنك تصدير نسخة احتياطية قبل الحذف.</p><div class="modal-actions"><button class="danger-button" data-action="confirm-delete-project">حذف المشروع</button><button class="text-button" data-action="cancel-delete">إبقاء</button></div></section></div>`;
  if (p && state.present !== null) out += presentHtml(p);
  if (state.annot) out += annotHtml();
  out += syncModalHtml();
  return out;
}

/* ================= الصفحة ================= */
function welcomeHtml() {
  const pr = state.progress;
  return `<section class="empty-layout animate-in delay-1"><div class="upload-zone empty-upload" id="dropzone">
      <div class="upload-illustration">${icon('presentation', 25)}</div><h2>ابدأ أول مشروع لك</h2>
      <p>أنشئ مشروعًا من شرائح تضيف إليها النصوص والصور والتعليقات ومؤشرات الخطورة،<br>ويحلّله النظام فورًا. أو استورد ملف PowerPoint موجودًا.</p>
      <button class="primary-button" data-action="new-project">${icon('plus', 17)}مشروع جديد</button>
      <button class="text-button" data-action="import">${icon('upload', 15)}استيراد ملف PowerPoint أو نسخة احتياطية</button>
      ${pr ? `<div class="progress-line">${icon('loader', 16, 'spin')}جارٍ استيراد ${h(pr.name)} · ${pr.index} من ${pr.total}</div>` : ''}</div>
    <div class="empty-aside">
      <div class="empty-note"><div class="empty-note-symbol">${icon('layers', 18)}</div><div><h3>مشروع كامل بشرائح</h3><p>أضف نصًا وجداول وصورًا وتعليقات ومؤشرات بمستوى خطورة ومسؤول وموعد.</p></div></div>
      <div class="empty-note"><div class="empty-note-symbol warm">${icon('barChart', 18)}</div><div><h3>تحليل مباشر</h3><p>تُجمَّع الحالات وتُرتَّب بالأولوية تلقائيًا، وتُحفظ لتعود إليها وتتابعها حتى الإغلاق.</p></div></div>
      <div class="method-card"><span class="method-index">خصوصية</span><p>كل المشاريع والصور تبقى في هذا المتصفح، ويمكنك أخذ نسخة احتياطية في أي وقت.</p></div></div></section>`;
}

function mainHtml(p) {
  const pr = state.progress;
  const banners = `${state.error ? `<div class="error-banner" role="alert">${icon('alert', 18)}<span>${h(state.error)}</span><button data-action="dismiss-error" aria-label="إغلاق رسالة الخطأ">${icon('x')}</button></div>` : ''}
    ${state.notice ? `<div class="notice-banner" role="status">${icon('check')}${h(state.notice)}<button aria-label="إغلاق التنبيه" data-action="dismiss-notice">${icon('x', 15)}</button></div>` : ''}`;
  let body;
  if (state.view === 'board') body = boardHtml();
  else if (!p) body = state.loadingSaved ? '<div class="loading-state"><span class="skeleton-block"></span><span class="skeleton-block short"></span>جارٍ استعادة مشاريعك…</div>' : welcomeHtml();
  else if (state.mode === 'edit') body = projectBarHtml(p) + editorHtml(p);
  else { const a = active(), s = activeSummary(); body = projectBarHtml(p) + summaryHtml(a, s) + insightsHtml(a, s) + workspaceHtml(p, a, s); }
  return `<main class="project-main mobile-wrap">${banners}${body}
    ${pr && (p || state.view === 'board') ? `<div class="processing-toast">${icon('loader', 17, 'spin')}جارٍ استيراد ${h(pr.name)} <span>${pr.index} / ${pr.total}</span></div>` : ''}
    ${state.undo ? `<div class="undo-toast" role="status">تم حذف العنصر<button data-action="undo">${icon('undo', 14)}تراجع</button></div>` : ''}
    <footer class="page-footer"><span>وَرَق <i>·</i> مشاريع شرائح ومراجعات محلية</span><span>التحليل التلقائي قواعد نصية تحتاج إلى مراجعتك؛ وتحدد أنت مستوى الخطورة يدويًا عند الحاجة.</span></footer></main>`;
}

function headerHtml() {
  return `<header class="topbar"><div class="topbar-inner container-1440">
    <div class="row-3"><button class="menu-button" data-action="toggle-sidebar" aria-label="قائمة المشاريع">${icon('menu', 20)}</button><div class="brand-mark">${icon('presentation', 21)}</div><div><div class="brand-name">وَرَق</div><div class="brand-caption">مساحة المشاريع والمراجعة</div></div></div>
    <div class="privacy-chip desktop-only">${icon('shield', 15)}<span>متزامنة مع حسابك الخاص</span></div>
    <div class="row-2">${syncButtonHtml()}<button class="add-file-button" data-action="new-project">${icon('plus', 17)}<span>مشروع جديد</span></button></div></div></header>`;
}

let lastAnimated = null;
function render() {
  if (!sync.session) { $app.innerHTML = gateHtml(); document.title = 'تسجيل الدخول'; return; }
  const el = document.activeElement;
  const keep = el && el.id && $app.contains(el) ? { id: el.id, start: el.selectionStart, end: el.selectionEnd } : null;
  const scrollY = window.scrollY;
  const p = currentProject();
  if (p) {
    if (ensureAllLayouts(p)) scheduleSave(p);
    if (state.sel && !curSlide()?.blocks.some(b => b.id === state.sel)) state.sel = null;
    if (state.editingId && !curSlide()?.blocks.some(b => b.id === state.editingId)) state.editingId = null;
  }
  const a = p && state.mode !== 'edit' ? active() : null;
  const s = a ? activeSummary() : null;
  const viewKey = `${state.view}:${state.projectId || 'empty'}:${state.mode}`;
  const settled = lastAnimated === viewKey;
  lastAnimated = viewKey;
  $app.innerHTML = `<div class="workspace-shell brand-orange${settled ? ' settled' : ''}" dir="rtl">${headerHtml()}
    <div class="app-body container-1440${state.sideCollapsed ? ' side-collapsed' : ''}">${sidebarHtml()}${state.sidebarOpen ? '<div class="sidebar-scrim" data-action="toggle-sidebar"></div>' : ''}${mainHtml(p)}</div>${state.sideCollapsed ? `<button class="side-fold side-toggle" data-action="toggle-side-collapse">${icon('menu', 14)} القائمة</button>` : ''}${modalsHtml(p, p ? active() : null, p ? activeSummary() : null)}</div>`;
  autosizeAll();
  setSaveIndicator();
  const target = state.focusId && document.getElementById(state.focusId);
  if (target) { target.focus({ preventScroll: false }); try { target.setSelectionRange(target.value.length, target.value.length); } catch { /* حقل بلا مؤشر */ } state.focusId = null; }
  else if (keep) {
    const next = document.getElementById(keep.id);
    if (next) { next.focus({ preventScroll: true }); try { next.setSelectionRange(keep.start, keep.end); } catch { /* حقول بلا مؤشر نصي */ } }
    window.scrollTo(0, scrollY);
  } else window.scrollTo(0, scrollY);
}
function renderCasesOnly() {
  const a = active(), holder = document.getElementById('cases-list');
  if (!a || !holder) return render();
  const s = activeSummary();
  holder.innerHTML = casesListHtml(a, s);
  const badge = document.getElementById('count-badge');
  if (badge) badge.textContent = `${casesView(a, s).length} حالة`;
}

/* ================= لوحة الشريحة (شبيهة بـ PowerPoint): عناصر حرة الموضع والحجم ================= */
const STAGE_AR = 16 / 9;
const COL_FULL = { x: 4, w: 92 }, COL_SIDE = { x: 50, w: 46 }, IMG_REGION = { x: 4, w: 42 };
const STAGE_TOP = 16, STAGE_BOTTOM = 97;
const isImg = (b) => b.type === 'image';
const selBlock = () => curSlide()?.blocks.find(b => b.id === state.sel) || null;

/* ---------- تقدير الارتفاعات (٪ من ارتفاع اللوحة) ---------- */
function textLines(text, fs, w, pad = 2.4) {
  const cpl = Math.max(8, Math.floor((w * 10 - pad * 10) / (fs * 10 * 0.52)));
  return String(text || ' ').split('\n').reduce((n, l) => n + Math.max(1, Math.ceil(l.length / cpl)), 0);
}
const lineH = (fs, lh = 1.6) => fs * 10 * lh / 5.625;
function estimateH(b, w) {
  if (b.type === 'text') { const fs = b.style === 'h' ? 2.3 : 1.7; return Math.max(5, textLines(b.text, fs, w) * lineH(fs, b.style === 'h' ? 1.4 : 1.6) + 2.5); }
  if (b.type === 'finding') return Math.max(10, lineH(1.2) + textLines(b.text, 1.5, w, 3) * lineH(1.5) + (b.cause ? textLines(b.cause, 1.35, w, 3) * lineH(1.35) : 0) + (b.action ? textLines(b.action, 1.35, w, 3) * lineH(1.35) : 0) + 4);
  if (b.type === 'comment') return Math.max(6, textLines(b.text, 1.4, w, 3) * lineH(1.4) + 3);
  if (b.type === 'table') { const cols = b.rows[0]?.length || 1; return b.rows.reduce((n, r) => n + Math.max(...r.map(c => textLines(c, 1.4, w / cols, 1.6))) * lineH(1.4, 1.35) + 1.8, 0) + 1; }
  return w * 0.75 * STAGE_AR;
}
function gridFor(n, region) {
  const cols = n <= 1 ? 1 : n <= 4 ? 2 : 3;
  const gap = 2, cellW = region.w === 92 && cols === 1 ? 60 : (region.w - gap * (cols - 1)) / cols;
  const offX = region.w === 92 && cols === 1 ? 20 : region.x;
  return { cols, gap, cellW, offX };
}
/** يوزّع كل عناصر الشريحة تلقائيًا (أول مرة): النصوص في عمود والصور في شبكة. */
function layoutAll(slide, items) {
  const imgs = items.filter(isImg), others = items.filter(b => !isImg(b));
  const col = imgs.length && others.length ? COL_SIDE : COL_FULL;
  let y = STAGE_TOP;
  for (const b of others) { const h = estimateH(b, col.w); b.box = { x: col.x, y, w: col.w, h, auto: true }; y += h + 1.6; }
  const region = others.length ? IMG_REGION : { x: 4, w: 92 };
  const g = gridFor(imgs.length, region);
  imgs.forEach((b, i) => {
    const w = g.cellW, h = w * 0.75 * STAGE_AR, row = Math.floor(i / g.cols), c = i % g.cols;
    b.box = { x: g.offX + c * (w + g.gap), y: STAGE_TOP + row * (h + g.gap), w, h, fit: true, auto: true };
    if (b.box.y + h > 99) b.box.y = Math.max(STAGE_TOP, 99 - h - (i % 4) * 2);
  });
}
/** يضع عنصرًا جديدًا في شريحة فيها عناصر موضوعة أصلًا. */
function placeNew(slide, b) {
  const others = slide.blocks.filter(x => x !== b && x.box);
  if (isImg(b)) {
    const imgs = others.filter(isImg), texts = others.filter(x => !isImg(x));
    if (!imgs.length && texts.length && texts.every(o => o.box.auto && o.box.w >= 90)) {
      let y = STAGE_TOP; for (const o of texts) { const h = estimateH(o, COL_SIDE.w); o.box = { x: COL_SIDE.x, y, w: COL_SIDE.w, h, auto: true }; y += h + 1.6; }
    }
    const region = texts.length ? IMG_REGION : { x: 4, w: 92 };
    const all = [...imgs, b], n = all.length;
    if (imgs.every(o => o.box.auto)) {           // لم يلمس المستخدم الصور: أعد توزيعها كلها في شبكة مرتبة
      const g = gridFor(n, region);
      all.forEach((o, i) => { const w = g.cellW, h = w * 0.75 * STAGE_AR, row = Math.floor(i / g.cols), c = i % g.cols; o.box = { x: g.offX + c * (w + g.gap), y: STAGE_TOP + row * (h + g.gap), w, h, fit: true, auto: true }; if (o.box.y + h > 99) o.box.y = Math.max(STAGE_TOP, 99 - h - (i % 4) * 2); });
    } else {                                      // الشبكة معدَّلة يدويًا: ضع الجديدة بحجم مناسب في مكان فارغ تقريبًا
      const w = Math.min(24, region.w), h = w * 0.75 * STAGE_AR, k = imgs.length;
      b.box = { x: 6 + (k % 6) * 3, y: STAGE_TOP + 2 + (k % 6) * 3, w, h, fit: true, auto: false };
    }
    return;
  }
  const texts = others.filter(x => !isImg(x));
  const hasImg = others.some(isImg);
  const col = texts.length ? { x: texts.at(-1).box.x, w: texts.at(-1).box.w } : hasImg ? COL_SIDE : COL_FULL;
  const h = estimateH(b, col.w);
  let y = texts.length ? Math.max(...texts.map(x => x.box.y + x.box.h)) + 1.6 : STAGE_TOP;
  let x = col.x;
  if (y + h > STAGE_BOTTOM) { y = STAGE_TOP + (others.length % 5) * 3; x = Math.min(100 - col.w - 1, col.x + (others.length % 5) * 2); }
  b.box = { x, y, w: col.w, h, auto: true };
}
function ensureAllLayouts(p) {
  let changed = false;
  for (const slide of p.slides) {
    const missing = slide.blocks.filter(b => !b.box);
    if (!missing.length) continue;
    changed = true;
    if (missing.length === slide.blocks.length) layoutAll(slide, missing); else missing.forEach(b => placeNew(slide, b));
  }
  return changed;
}

/* ---------- العرض ---------- */
const sevOf = (b) => SEVERITY_CHOICES.includes(b.severity) ? b.severity : severityFor(cleanText(`${b.text} ${b.cause}`), 'مشكلة');
function objInner(b, edit) {
  if (b.type === 'text') {
    if (state.editingId === b.id && edit) return `<textarea class="inline-edit st-${b.style}" id="inl-${b.id}" data-block="${b.id}" data-field="text" aria-label="نص">${h(b.text)}</textarea>`;
    if (!b.text.trim()) return edit ? '<span class="ph">انقر مرتين للكتابة</span>' : '';
    if (b.style === 'h') return `<div class="st-h">${h(b.text)}</div>`;
    if (b.style === 'ul') return `<ul class="st-ul">${b.text.split('\n').filter(l => l.trim()).map(l => `<li>${h(l)}</li>`).join('')}</ul>`;
    return `<div class="st-p">${h(b.text).replace(/\n/g, '<br>')}</div>`;
  }
  if (b.type === 'finding') {
    const sev = sevOf(b);
    if (!`${b.text}${b.cause}${b.action}`.trim()) return edit ? '<span class="ph">مؤشر فارغ — حرّره من اللوحة أدناه</span>' : '';
    return `<div class="sf-meta"><span class="severity-tag ${severityTone[sev]}">${sev}</span><span>${h(STATUS_LABELS[b.status] || '')}</span>${b.owner ? `<span>${icon('user', 11)}${h(b.owner)}</span>` : ''}${b.due ? `<span>${icon('calendar', 11)}${h(b.due)}</span>` : ''}</div>
      ${b.text ? `<div class="sf-text">${h(b.text)}</div>` : ''}${b.cause ? `<div class="sf-line"><em>السبب المحتمل:</em> ${h(b.cause)}</div>` : ''}${b.action ? `<div class="sf-line sf-action"><em>الإجراء المقترح:</em> ${h(b.action)}</div>` : ''}`;
  }
  if (b.type === 'image') {
    const src = state.assetUrls[b.assetId];
    return `${src ? `<img class="stage-img" data-img="${b.id}" src="${h(src)}" alt="${h(b.caption || 'صورة')}" draggable="false">` : '<div class="so-loading">جارٍ التحميل…</div>'}${b.caption ? `<div class="cap">${h(b.caption)}</div>` : ''}`;
  }
  if (b.type === 'comment') return `<div class="sticky-head">${icon('message', 12)}تعليق${b.author ? ` · ${h(b.author)}` : ''}</div>${state.editingId === b.id && edit ? `<textarea class="inline-edit" id="inl-${b.id}" data-block="${b.id}" data-field="text">${h(b.text)}</textarea>` : (b.text.trim() ? h(b.text).replace(/\n/g, '<br>') : '<span class="ph">انقر مرتين للكتابة</span>')}`;
  if (b.type === 'table') return `<table class="st-table"><tbody>${b.rows.map((r, i) => `<tr>${r.map(c => i === 0 ? `<th>${h(c)}</th>` : `<td>${h(c)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  return '';
}
const HANDLES_ALL = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'], HANDLES_IMG = ['nw', 'ne', 'se', 'sw'];
const handlesHtml = (b) => `<span class="handles">${(isImg(b) ? HANDLES_IMG : HANDLES_ALL).map(d => `<span class="handle h-${d}" data-dir="${d}"></span>`).join('')}</span>`;
function boxStyle(b) { const x = b.box; return `left:${x.x}%;top:${x.y}%;width:${x.w}%;${isImg(b) ? `height:${x.h}%` : `min-height:${x.h}%`}`; }
function objHtml(b, edit) {
  const sel = edit && state.sel === b.id;
  const sev = b.type === 'finding' ? ` sev-${sevOf(b)}` : '';
  return `<div class="stage-obj so-${b.type}${sev}${sel ? ' selected' : ''}${state.editingId === b.id && edit ? ' editing' : ''}" data-obj="${b.id}" style="${boxStyle(b)}"><div class="so-inner">${objInner(b, edit)}</div>${sel ? handlesHtml(b) : ''}</div>`;
}
function stageHtml(slide, edit) {
  const objs = slide.blocks.filter(b => b.box && (edit || b.type !== 'comment')).map(b => objHtml(b, edit)).join('');
  return `<div class="stage${edit ? '' : ' static'}" ${edit ? 'id="stage"' : ''} style="--zoom:${state.zoom}"><div class="stage-title"${edit ? ' data-title-obj' : ''}>${h(slide.title || '')}${edit && !slide.title ? '<span class="ph">انقر مرتين لكتابة عنوان الشريحة</span>' : ''}</div>${objs}<div class="guide gv" hidden></div><div class="guide gh" hidden></div></div>`;
}
function miniHtml(slide) {
  const objs = slide.blocks.filter(b => b.box && b.type !== 'comment').map(b => {
    const x = b.box, st = `left:${x.x}%;top:${x.y}%;width:${x.w}%;height:${x.h}%`;
    if (isImg(b)) { const src = state.assetUrls[b.assetId]; return `<i class="m-obj m-image" style="${st}">${src ? `<img src="${h(src)}" alt="">` : ''}</i>`; }
    return `<i class="m-obj m-${b.type}${b.type === 'finding' ? ` m-${sevOf(b) === 'مرتفع' ? 'hi' : sevOf(b) === 'متوسط' ? 'mid' : 'lo'}` : ''}" style="${st}"></i>`;
  }).join('');
  return `<span class="mini-stage"><i class="m-title"></i>${objs}</span>`;
}

function ribbonInner() {
  const slide = curSlide(), b = selBlock();
  const rb = (act, ico, label, extra = '', off = false) => `<button class="rb" data-action="${act}" ${extra} title="${label}" aria-label="${label}"${off ? ' disabled' : ''}>${icon(ico, 16)}<span>${label}</span></button>`;
  const h0 = state.hist;
  return `<div class="rb-group">${rb('hist-undo', 'undo', 'تراجع', '', !h0.past.length)}${rb('hist-redo', 'redo', 'إعادة', '', !h0.future.length)}</div>
    <div class="rb-group"><span class="rb-label">إدراج</span>${rb('add-block', 'type', 'نص', 'data-type="text"')}${rb('add-image', 'imagePlus', 'صورة')}${rb('add-block', 'flag', 'مؤشر', 'data-type="finding"')}${rb('add-block', 'table', 'جدول', 'data-type="table"')}${rb('add-block', 'message', 'تعليق', 'data-type="comment"')}</div>
    <div class="rb-group"><span class="rb-label">العنصر</span>${rb('annotate', 'pen', 'تعليم وقص الصورة', `data-id="${b?.id || ''}"`, !(b && isImg(b)))}${rb('dup-block', 'copy', 'نسخ', `data-id="${b?.id || ''}"`, !b)}${rb('del-block', 'trash', 'حذف', `data-id="${b?.id || ''}"`, !b)}${rb('obj-front', 'layers', 'للأمام', '', !b)}${rb('obj-back', 'layers', 'للخلف', '', !b)}${rb('obj-center-h', 'alignH', 'توسيط أفقي', '', !b)}${rb('obj-center-v', 'alignV', 'توسيط رأسي', '', !b)}${rb('obj-fill', 'maximize', 'ملء الشريحة', '', !b || !isImg(b))}</div>
    <div class="rb-group zoom-group">${rb('zoom-out', 'zoomOut', 'تصغير')}<span class="zoom-val" id="zoom-val">${Math.round(state.zoom * 100)}%</span>${rb('zoom-in', 'zoomIn', 'تكبير')}${rb('zoom-fit', 'rotateCcw', '100%')}</div>`;
}
function inspectorInner(b) {
  if (!b) return `<div class="inspector-hint">${icon('pointer', 18)}<p>اضغط على أي عنصر في الشريحة لتحريكه أو تغيير حجمه من المقابض. انقر مرتين على نص لتكتب فيه مباشرة، وعلى صورة لتعليمها.<br>لإضافة عدة صور: زر «صورة» أو الصق/اسحب الصور فوق الشريحة.</p></div>`;
  const slide = curSlide(); const i = slide.blocks.indexOf(b);
  return blockHtml(b, i, slide.blocks.length);
}
function stageEditorBody(p, slide) {
  const b = selBlock();
  return `<div class="ribbon" id="ribbon">${ribbonInner()}</div>
    <div class="stage-wrap" id="stage-wrap">${stageHtml(slide, true)}</div>
    <div class="inspector" id="inspector">${inspectorInner(b)}</div>`;
}

/* ---------- التحديد بلا إعادة رسم كاملة ---------- */
function refreshSelection() {
  document.querySelectorAll('.stage-obj.selected').forEach(el => { el.classList.remove('selected'); el.querySelector('.handles')?.remove(); });
  const b = selBlock();
  if (b) { const el = document.querySelector(`[data-obj="${b.id}"]`); if (el) { el.classList.add('selected'); el.insertAdjacentHTML('beforeend', handlesHtml(b)); } }
  const ins = document.getElementById('inspector'); if (ins) ins.innerHTML = inspectorInner(b);
  const rb = document.getElementById('ribbon'); if (rb) rb.innerHTML = ribbonInner();
  autosizeAll();
}
function refreshObj(id) {
  const b = curSlide()?.blocks.find(x => x.id === id), el = document.querySelector(`[data-obj="${id}"]`);
  if (!b || !el) return;
  el.querySelector('.so-inner').innerHTML = objInner(b, true);
  el.classList.toggle('editing', state.editingId === id);
}

/* ---------- التراجع والإعادة (لقطات لكل التعديلات) ---------- */
const snapNow = () => { const p = currentProject(); return p ? JSON.stringify({ slides: p.slides, imageAssets: p.imageAssets }) : ''; };
function checkpoint() {
  const s = snapNow(); if (!s) return;
  if (state.hist.past.at(-1) === s) return;
  state.hist.past.push(s); if (state.hist.past.length > 80) state.hist.past.shift();
  state.hist.future = [];
}
function restoreSnap(s) {
  const p = currentProject(), d = JSON.parse(s);
  p.slides = d.slides; p.imageAssets = d.imageAssets;
  state.slideIdx = Math.min(state.slideIdx, p.slides.length - 1); state.sel = null; state.editingId = null;
  touch(p); render();
}
function histUndo() { if (!state.hist.past.length) return; state.hist.future.push(snapNow()); restoreSnap(state.hist.past.pop()); }
function histRedo() { if (!state.hist.future.length) return; state.hist.past.push(snapNow()); restoreSnap(state.hist.future.pop()); }

/* ---------- السحب وتغيير الحجم بالمؤشر ---------- */
let gesture = null;
const stageEl = () => document.getElementById('stage');
function applyBox(el, b) { el.style.left = `${b.box.x}%`; el.style.top = `${b.box.y}%`; el.style.width = `${b.box.w}%`; if (isImg(b)) el.style.height = `${b.box.h}%`; else el.style.minHeight = `${b.box.h}%`; }
document.addEventListener('pointerdown', (e) => {
  if (state.mode !== 'edit' || state.editView !== 'stage' || e.button > 0) return;
  const stage = stageEl(); if (!stage || !stage.contains(e.target) || e.target.closest('.inline-edit')) return;
  const obj = e.target.closest('[data-obj]');
  if (!obj) { if (state.sel) { state.sel = null; refreshSelection(); } return; }
  const id = obj.dataset.obj, handle = e.target.closest('.handle');
  if (state.sel !== id) { state.sel = id; refreshSelection(); }
  const b = curSlide().blocks.find(x => x.id === id); if (!b) return;
  const img = obj.querySelector('.stage-img');
  gesture = { kind: handle ? 'resize' : 'move', dir: handle?.dataset.dir, b, sx: e.clientX, sy: e.clientY, box: { ...b.box }, rect: stage.getBoundingClientRect(), moved: false, el: stage.querySelector(`[data-obj="${id}"]`),
    hw: isImg(b) && img && img.naturalWidth ? (img.naturalHeight / img.naturalWidth) * STAGE_AR : (b.box.h / b.box.w) };
  e.preventDefault();
});
document.addEventListener('pointermove', (e) => {
  const g = gesture; if (!g) return;
  const dxp = e.clientX - g.sx, dyp = e.clientY - g.sy;
  if (!g.moved && Math.hypot(dxp, dyp) < 4) return;
  if (!g.moved) { checkpoint(); g.moved = true; g.b.box.auto = false; g.b.box.fit = false; document.body.classList.add('gesturing'); }
  const dx = dxp / g.rect.width * 100, dy = dyp / g.rect.height * 100, o = g.box, bx = g.b.box;
  const vg = document.querySelector('.guide.gv'), hg = document.querySelector('.guide.gh');
  if (g.kind === 'move') {
    let x = Math.min(100 - o.w, Math.max(0, o.x + dx)), y = Math.min(100 - Math.min(o.h, 100), Math.max(0, o.y + dy));
    const snapV = Math.abs(x + o.w / 2 - 50) < 1.2, snapH = Math.abs(y + o.h / 2 - 50) < 1.2;
    if (snapV) x = 50 - o.w / 2; if (snapH) y = 50 - o.h / 2;
    if (vg) vg.hidden = !snapV; if (hg) hg.hidden = !snapH;
    bx.x = x; bx.y = y;
  } else {
    const d = g.dir, min = 4;
    let w = o.w, h = o.h, x = o.x, y = o.y;
    if (d.includes('e')) w = o.w + dx; if (d.includes('w')) w = o.w - dx;
    if (isImg(g.b)) { w = Math.max(min, w); h = w * g.hw; }
    else { w = Math.max(min, w); if (d.includes('s')) h = o.h + dy; if (d.includes('n')) h = o.h - dy; h = Math.max(3, h); }
    if (d.includes('w')) x = o.x + o.w - w; if (d.includes('n')) y = o.y + o.h - h;
    bx.x = x; bx.y = y; bx.w = w; bx.h = h;
  }
  applyBox(g.el, g.b);
});
document.addEventListener('pointerup', () => {
  const g = gesture; gesture = null;
  document.body.classList.remove('gesturing');
  document.querySelectorAll('.guide').forEach(x => { x.hidden = true; });
  if (g?.moved) { const p = currentProject(); touch(p); const ribbon = document.getElementById('ribbon'); if (ribbon) ribbon.innerHTML = ribbonInner(); }
});
/** الصورة تأخذ نسبتها الحقيقية أول ما تُحمَّل (إن لم يغيّر المستخدم حجمها). */
document.addEventListener('load', (e) => {
  const img = e.target; if (img?.tagName !== 'IMG' || !img.dataset.img || !img.naturalWidth) return;
  const b = curSlide()?.blocks.find(x => x.id === img.dataset.img);
  if (!b?.box?.fit) return;
  const hh = b.box.w * (img.naturalHeight / img.naturalWidth) * STAGE_AR;
  let w = b.box.w, h2 = hh; if (h2 > 82) { w = b.box.w * 82 / h2; h2 = 82; }
  b.box.w = w; b.box.h = h2; b.box.fit = false;
  const el = document.querySelector(`[data-obj="${b.id}"]`); if (el) applyBox(el, b);
  touch(currentProject());
}, true);

/* ---------- الكتابة المباشرة داخل العنصر ---------- */
function startInline(id) {
  const b = curSlide()?.blocks.find(x => x.id === id); if (!b || (b.type !== 'text' && b.type !== 'comment')) return;
  state.editingId = id; refreshObj(id);
  const t = document.getElementById(`inl-${id}`);
  if (t) { t.style.height = `${Math.max(t.closest('.stage-obj').offsetHeight, t.scrollHeight)}px`; t.focus(); t.setSelectionRange(t.value.length, t.value.length); }
}
document.addEventListener('dblclick', (e) => {
  if (state.mode !== 'edit' || state.editView !== 'stage') return;
  if (e.target.closest('[data-title-obj]')) { const t = document.getElementById('slide-title'); t?.scrollIntoView({ block: 'center' }); t?.focus(); return; }
  const obj = e.target.closest('[data-obj]'); if (!obj) return;
  const b = curSlide().blocks.find(x => x.id === obj.dataset.obj); if (!b) return;
  if (b.type === 'text' || b.type === 'comment') startInline(b.id);
  else if (b.type === 'image') actions['annotate']({ dataset: { id: b.id } });
  else document.querySelector('#inspector textarea, #inspector input')?.focus();
});
document.addEventListener('focusout', (e) => {
  if (!e.target.classList?.contains('inline-edit')) return;
  const id = e.target.dataset.block;
  setTimeout(() => { if (state.editingId === id && document.activeElement?.id !== `inl-${id}`) { state.editingId = null; refreshObj(id); const ins = document.getElementById('inspector'); if (ins && state.sel === id) ins.innerHTML = inspectorInner(selBlock()); } }, 0);
});
document.addEventListener('focusin', (e) => { if (e.target.dataset?.block || e.target.dataset?.slideTitle || e.target.dataset?.slideNotes) checkpoint(); });

/* ---------- لوحة المفاتيح ---------- */
function stageKey(e) {
  if (state.mode !== 'edit' || state.annot) return false;
  const tag = e.target.tagName, typing = tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT';
  if ((e.ctrlKey || e.metaKey) && !typing && e.key.toLowerCase() === 'z') { e.preventDefault(); e.shiftKey ? histRedo() : histUndo(); return true; }
  if ((e.ctrlKey || e.metaKey) && !typing && e.key.toLowerCase() === 'y') { e.preventDefault(); histRedo(); return true; }
  if (state.editView !== 'stage' || typing) return false;
  const b = selBlock(); if (!b) return false;
  if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); actions['del-block']({ dataset: { id: b.id } }); return true; }
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'd') { e.preventDefault(); actions['dup-block']({ dataset: { id: b.id } }); return true; }
  if (e.key === 'Escape') { state.sel = null; refreshSelection(); return true; }
  if (e.key === 'Enter' && (b.type === 'text' || b.type === 'comment')) { e.preventDefault(); startInline(b.id); return true; }
  const step = e.shiftKey ? 5 : 1, d = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[e.key];
  if (d) {
    e.preventDefault(); checkpoint();
    b.box.x = Math.min(100 - b.box.w, Math.max(0, b.box.x + d[0])); b.box.y = Math.min(99, Math.max(0, b.box.y + d[1])); b.box.auto = false;
    const el = document.querySelector(`[data-obj="${b.id}"]`); if (el) applyBox(el, b); touch(currentProject()); return true;
  }
  return false;
}
document.addEventListener('wheel', (e) => {
  if (!(e.ctrlKey || e.metaKey) || !e.target.closest?.('#stage-wrap')) return;
  e.preventDefault(); setZoom(state.zoom + (e.deltaY < 0 ? 0.1 : -0.1));
}, { passive: false });
function setZoom(z) {
  state.zoom = Math.min(2.5, Math.max(0.4, Math.round(z * 10) / 10));
  const st = stageEl(); if (st) st.style.setProperty('--zoom', state.zoom);
  const v = document.getElementById('zoom-val'); if (v) v.textContent = `${Math.round(state.zoom * 100)}%`;
}

const stageActions = {
  'hist-undo': () => histUndo(), 'hist-redo': () => histRedo(),
  'zoom-in': () => setZoom(state.zoom + 0.1), 'zoom-out': () => setZoom(state.zoom - 0.1), 'zoom-fit': () => setZoom(1),
  'edit-view': (t) => { state.editView = t.dataset.view; state.sel = null; state.editingId = null; render(); },
  'obj-front': () => { const b = selBlock(); if (!b) return; mutate(() => { const a = curSlide().blocks; a.splice(a.indexOf(b), 1); a.push(b); }); },
  'obj-back': () => { const b = selBlock(); if (!b) return; mutate(() => { const a = curSlide().blocks; a.splice(a.indexOf(b), 1); a.unshift(b); }); },
  'obj-center-h': () => { const b = selBlock(); if (b) mutate(() => { b.box.x = 50 - b.box.w / 2; b.box.auto = false; }); },
  'obj-center-v': () => { const b = selBlock(); if (b) mutate(() => { b.box.y = Math.max(0, 50 - b.box.h / 2); b.box.auto = false; }); },
  'obj-fill': () => { const b = selBlock(); if (b && isImg(b)) mutate(() => { b.box = { x: 4, y: STAGE_TOP - 1, w: 92, h: 92 * 0.75 * STAGE_AR, fit: true }; }); },
};

/* ================= تعليم الصور: أسهم، أشكال، قلم، تظليل، نص، ترقيم، قص، تكبير ================= */
const ANN_COLORS = ['#e53935', '#e96826', '#fbc02d', '#43a047', '#1e88e5', '#212121', '#ffffff'];
const ANN_WIDTHS = [3, 6, 11], ANN_FONTS = [26, 38, 56];
const ANN_TOOLS = [['select', 'pointer', 'تحديد'], ['arrow', 'arrowUpRight', 'سهم'], ['rect', 'square', 'مستطيل'], ['ellipse', 'circle', 'دائرة'], ['pen', 'pen', 'قلم'], ['hi', 'highlighter', 'تظليل'], ['text', 'type', 'نص'], ['badge', 'hash', 'ترقيم'], ['crop', 'crop', 'قص']];
const esc2 = (s) => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);

/** عنصر SVG واحد لكل علامة؛ الإحداثيات مُطبَّعة (0..1) وتُرسم على لوحة عرضها 1000 وارتفاعها H. */
function markSvg(m, H, live = false) {
  const X = (v) => (v * 1000).toFixed(1), Y = (v) => (v * H).toFixed(1), sw = ANN_WIDTHS[m.s ?? 1], col = m.c || '#e53935';
  let body = '', hit = '';
  const stroke = `fill="none" stroke="${col}" stroke-width="${sw}" stroke-linecap="round" stroke-linejoin="round"`;
  const hitStroke = 'fill="none" stroke="rgba(0,0,0,0.001)" stroke-width="22" pointer-events="stroke"';
  if (m.t === 'rect') { const a = `x="${X(m.x)}" y="${Y(m.y)}" width="${(m.w * 1000).toFixed(1)}" height="${(m.h * H).toFixed(1)}"`; body = `<rect ${a} rx="3" ${stroke}/>`; hit = `<rect ${a} ${hitStroke}/>`; }
  else if (m.t === 'ellipse') { const a = `cx="${X(m.x + m.w / 2)}" cy="${Y(m.y + m.h / 2)}" rx="${(m.w * 500).toFixed(1)}" ry="${(m.h * H / 2).toFixed(1)}"`; body = `<ellipse ${a} ${stroke}/>`; hit = `<ellipse ${a} ${hitStroke}/>`; }
  else if (m.t === 'hi') { const a = `x="${X(m.x)}" y="${Y(m.y)}" width="${(m.w * 1000).toFixed(1)}" height="${(m.h * H).toFixed(1)}"`; body = `<rect ${a} fill="${col}" fill-opacity="0.38"/>`; hit = `<rect ${a} fill="rgba(0,0,0,0.001)" pointer-events="all"/>`; }
  else if (m.t === 'arrow') {
    const x1 = m.x1 * 1000, y1 = m.y1 * H, x2 = m.x2 * 1000, y2 = m.y2 * H, ang = Math.atan2(y2 - y1, x2 - x1), len = sw * 3.6 + 12, wing = 0.46;
    const p1 = [x2 - len * Math.cos(ang - wing), y2 - len * Math.sin(ang - wing)], p2 = [x2 - len * Math.cos(ang + wing), y2 - len * Math.sin(ang + wing)];
    body = `<line x1="${x1.toFixed(1)}" y1="${y1.toFixed(1)}" x2="${(x2 - Math.cos(ang) * len * 0.6).toFixed(1)}" y2="${(y2 - Math.sin(ang) * len * 0.6).toFixed(1)}" ${stroke}/><polygon points="${x2.toFixed(1)},${y2.toFixed(1)} ${p1[0].toFixed(1)},${p1[1].toFixed(1)} ${p2[0].toFixed(1)},${p2[1].toFixed(1)}" fill="${col}"/>`;
    hit = `<line x1="${x1.toFixed(1)}" y1="${y1.toFixed(1)}" x2="${x2.toFixed(1)}" y2="${y2.toFixed(1)}" ${hitStroke}/>`;
  } else if (m.t === 'pen') { const pts = m.pts.map(p => `${X(p[0])},${Y(p[1])}`).join(' '); body = `<polyline points="${pts}" ${stroke}/>`; hit = `<polyline points="${pts}" ${hitStroke}/>`; }
  else if (m.t === 'text') { const fs = ANN_FONTS[m.s ?? 1]; body = `<text x="${X(m.x)}" y="${Y(m.y)}" font-size="${fs}" font-weight="700" font-family="Arial, Tahoma, sans-serif" fill="${col}" stroke="${col === '#ffffff' ? '#222' : '#fff'}" stroke-width="${fs / 9}" paint-order="stroke" direction="rtl" text-anchor="start" dominant-baseline="hanging">${esc2(m.text)}</text>`; hit = `<rect x="${X(m.x)}" y="${Y(m.y)}" width="${Math.max(60, m.text.length * fs * 0.55)}" height="${fs * 1.3}" fill="rgba(0,0,0,0.001)" pointer-events="all"/>`; }
  else if (m.t === 'badge') { const r = 15 + sw * 2; body = `<circle cx="${X(m.x)}" cy="${Y(m.y)}" r="${r}" fill="${col}" stroke="#fff" stroke-width="3"/><text x="${X(m.x)}" y="${Y(m.y)}" font-size="${r * 1.25}" font-weight="700" font-family="Arial, sans-serif" fill="${col === '#ffffff' || col === '#fbc02d' ? '#222' : '#fff'}" text-anchor="middle" dominant-baseline="central">${m.n}</text>`; hit = `<circle cx="${X(m.x)}" cy="${Y(m.y)}" r="${r + 6}" fill="rgba(0,0,0,0.001)" pointer-events="all"/>`; }
  return live ? `<g data-mark="${m.id}" class="ann-mark">${body}${hit}</g>` : `<g>${body}</g>`;
}
function markBBox(m) {
  if (m.t === 'rect' || m.t === 'ellipse' || m.t === 'hi') return { x: m.x, y: m.y, w: m.w, h: m.h };
  if (m.t === 'arrow') return { x: Math.min(m.x1, m.x2), y: Math.min(m.y1, m.y2), w: Math.abs(m.x2 - m.x1), h: Math.abs(m.y2 - m.y1) };
  if (m.t === 'pen') { const xs = m.pts.map(p => p[0]), ys = m.pts.map(p => p[1]); return { x: Math.min(...xs), y: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys) }; }
  return { x: m.x - 0.02, y: m.y - 0.02, w: 0.08, h: 0.06 };
}
function moveMark(m, dx, dy) {
  if ('x' in m) { m.x += dx; m.y += dy; }
  if (m.t === 'arrow') { m.x1 += dx; m.x2 += dx; m.y1 += dy; m.y2 += dy; }
  if (m.t === 'pen') m.pts = m.pts.map(p => [p[0] + dx, p[1] + dy]);
}
function remapMark(m, c) {
  const fx = (v) => (v - c.x) / c.w, fy = (v) => (v - c.y) / c.h;
  if ('x' in m) { m.x = fx(m.x); m.y = fy(m.y); if ('w' in m) { m.w /= c.w; m.h /= c.h; } }
  if (m.t === 'arrow') { m.x1 = fx(m.x1); m.x2 = fx(m.x2); m.y1 = fy(m.y1); m.y2 = fy(m.y2); }
  if (m.t === 'pen') m.pts = m.pts.map(p => [fx(p[0]), fy(p[1])]);
}

/* ---------- الدمج في صورة جديدة ---------- */
async function flattenMarks(blob, marks, nw, nh) {
  const bmp = await createImageBitmap(blob);
  const k = Math.min(1, 2000 / Math.max(nw, nh)), W = Math.round(nw * k), Hh = Math.round(nh * k);
  const canvas = document.createElement('canvas'); canvas.width = W; canvas.height = Hh;
  const ctx = canvas.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, Hh); ctx.drawImage(bmp, 0, 0, W, Hh); bmp.close?.();
  const H = 1000 * nh / nw;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 ${H}" width="${W}" height="${Hh}">${marks.map(m => markSvg(m, H)).join('')}</svg>`;
  const url = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml' }));
  const img = new Image();
  await new Promise((res, rej) => { img.onload = res; img.onerror = rej; img.src = url; });
  ctx.drawImage(img, 0, 0, W, Hh); URL.revokeObjectURL(url);
  const type = blob.type === 'image/png' && W * Hh < 1.5e6 ? 'image/png' : 'image/jpeg';
  return new Promise(r => canvas.toBlob(r, type, 0.9));
}

/* ---------- الحالة والعرض ---------- */
const annSnap = () => { const a = state.annot; return { marks: JSON.stringify(a.marks), srcBlob: a.srcBlob, srcUrl: a.srcUrl, nw: a.nw, nh: a.nh, cropped: a.cropped }; };
function annPush() { const a = state.annot; a.past.push(annSnap()); if (a.past.length > 60) a.past.shift(); a.future = []; a.dirty = true; }
function annRestore(s) { const a = state.annot; a.marks = JSON.parse(s.marks); a.srcBlob = s.srcBlob; a.srcUrl = s.srcUrl; a.nw = s.nw; a.nh = s.nh; a.cropped = s.cropped; a.sel = null; a.crop = null; renderAnnot(); }
function annUndo() { const a = state.annot; if (!a?.past.length) return; a.future.push(annSnap()); annRestore(a.past.pop()); }
function annRedo() { const a = state.annot; if (!a?.future.length) return; a.past.push(annSnap()); annRestore(a.future.pop()); }
const annH = () => 1000 * state.annot.nh / state.annot.nw;

function annotHtml() {
  const a = state.annot; if (!a) return '';
  const H = annH();
  const tool = (t) => `<button class="at${a.tool === t[0] ? ' on' : ''}" data-action="ann-tool" data-tool="${t[0]}" title="${t[2]}" aria-label="${t[2]}">${icon(t[1], 16)}<span>${t[2]}</span></button>`;
  const cropBox = a.crop ? `<path class="crop-mask" fill-rule="evenodd" d="M0 0H1000V${H}H0Z M${a.crop.x * 1000} ${a.crop.y * H}h${a.crop.w * 1000}v${a.crop.h * H}h${-a.crop.w * 1000}Z"/><rect class="crop-frame" x="${a.crop.x * 1000}" y="${a.crop.y * H}" width="${a.crop.w * 1000}" height="${a.crop.h * H}"/>` : '';
  return `<div class="modal-backdrop annot-backdrop"><section class="annot-modal" role="dialog" aria-modal="true" aria-label="تعليم الصورة">
    <div class="annot-bar"><strong class="annot-title">${icon('pen', 16)}تعليم الصورة</strong>
      <div class="at-group">${ANN_TOOLS.map(tool).join('')}</div>
      <div class="at-group"><span class="at-label">اللون</span>${ANN_COLORS.map(c => `<button class="swatch${a.color === c ? ' on' : ''}" style="background:${c}" data-action="ann-color" data-color="${c}" aria-label="لون ${c}"></button>`).join('')}</div>
      <div class="at-group"><span class="at-label">السُّمك</span>${['رفيع', 'متوسط', 'سميك'].map((l, i) => `<button class="at sz${a.size === i ? ' on' : ''}" data-action="ann-size" data-size="${i}">${l}</button>`).join('')}</div>
      <div class="at-group"><button class="at" data-action="ann-undo" ${a.past.length ? '' : 'disabled'} aria-label="تراجع">${icon('undo', 15)}<span>تراجع</span></button><button class="at" data-action="ann-redo" ${a.future.length ? '' : 'disabled'} aria-label="إعادة">${icon('redo', 15)}<span>إعادة</span></button>
        <button class="at" data-action="ann-delete" ${a.sel ? '' : 'disabled'}>${icon('trash', 15)}<span>حذف المحدد</span></button><button class="at" data-action="ann-clear" ${a.marks.length ? '' : 'disabled'}>مسح الكل</button></div>
      <div class="at-group"><button class="at" data-action="ann-zoom-out" aria-label="تصغير">${icon('zoomOut', 15)}</button><span class="zoom-val">${Math.round(a.zoom * 100)}%</span><button class="at" data-action="ann-zoom-in" aria-label="تكبير">${icon('zoomIn', 15)}</button><button class="at" data-action="ann-zoom-fit">100%</button></div>
      ${a.crop ? `<div class="at-group crop-actions"><button class="primary-button sm" data-action="ann-crop-apply">${icon('check', 15)}تطبيق القص</button><button class="text-button" data-action="ann-crop-cancel">إلغاء القص</button></div>` : ''}
      <div class="at-group end"><button class="primary-button" data-action="ann-save" id="ann-save">${icon('save', 16)}حفظ التعليم</button><button class="text-button" data-action="ann-cancel">${a.confirmCancel ? 'تأكيد التجاهل' : 'إلغاء'}</button></div></div>
    <div class="annot-hint">${a.tool === 'text' ? 'انقر على الصورة ثم اكتب النص واضغط Enter.' : a.tool === 'crop' ? 'اسحب لتحديد المنطقة المراد الإبقاء عليها ثم اضغط «تطبيق القص».' : a.tool === 'select' ? 'انقر على علامة لتحديدها وتحريكها، وDelete لحذفها.' : a.tool === 'badge' ? 'انقر لوضع رقم متسلسل على الصورة.' : 'اسحب فوق الصورة للرسم. Ctrl+عجلة الفأرة للتكبير.'}</div>
    <div class="annot-scroll" id="annot-scroll"><div class="annot-canvas tool-${a.tool}" id="annot-canvas" style="--ar:${a.nw / a.nh};--z:${a.zoom}">
      <img src="${esc2(a.srcUrl)}" alt="الصورة الأصلية" draggable="false">
      <svg id="ann-svg" viewBox="0 0 1000 ${H}" preserveAspectRatio="none"><g id="ann-marks">${a.marks.map(m => markSvg(m, H, true)).join('')}</g><g id="ann-sel"></g><g id="ann-draft"></g>${cropBox}</svg>
      ${a.textAt ? `<input class="ann-text-input" id="ann-text" style="left:${a.textAt.x * 100}%;top:${a.textAt.y * 100}%" placeholder="اكتب النص ثم Enter" dir="rtl">` : ''}</div></div></section></div>`;
}
function renderAnnot() {
  const old = document.querySelector('.annot-backdrop'); if (!old) return render();
  const sc = document.getElementById('annot-scroll'), pos = sc ? [sc.scrollLeft, sc.scrollTop] : [0, 0];
  old.outerHTML = annotHtml();
  const n = document.getElementById('annot-scroll'); if (n) { n.scrollLeft = pos[0]; n.scrollTop = pos[1]; }
  drawSel(); document.getElementById('ann-text')?.focus();
}
function redrawMarks() {
  const a = state.annot, H = annH();
  const g = document.getElementById('ann-marks'); if (g) g.innerHTML = a.marks.map(m => markSvg(m, H, true)).join('');
  drawSel();
  const undo = document.querySelector('[data-action=ann-undo]'), redo = document.querySelector('[data-action=ann-redo]'), del = document.querySelector('[data-action=ann-delete]'), clr = document.querySelector('[data-action=ann-clear]');
  if (undo) undo.disabled = !a.past.length; if (redo) redo.disabled = !a.future.length; if (del) del.disabled = !a.sel; if (clr) clr.disabled = !a.marks.length;
}
function drawSel() {
  const a = state.annot, el = document.getElementById('ann-sel'); if (!a || !el) return;
  const m = a.sel && a.marks.find(x => x.id === a.sel);
  if (!m) { el.innerHTML = ''; return; }
  const b = markBBox(m), H = annH(), pad = 8;
  el.innerHTML = `<rect class="sel-box" x="${b.x * 1000 - pad}" y="${b.y * H - pad}" width="${b.w * 1000 + pad * 2}" height="${b.h * H + pad * 2}"/>`;
}

/* ---------- الرسم بالمؤشر ---------- */
let annG = null;
const annPt = (e) => { const r = document.getElementById('ann-svg').getBoundingClientRect(); return { x: Math.min(1, Math.max(0, (e.clientX - r.left) / r.width)), y: Math.min(1, Math.max(0, (e.clientY - r.top) / r.height)) }; };
function draftOf(g) {
  const a = state.annot, s = g.start, c = g.cur, base = { id: uid(), t: g.t, c: a.color, s: a.size };
  if (g.t === 'arrow') return { ...base, x1: s.x, y1: s.y, x2: c.x, y2: c.y };
  if (g.t === 'pen') return { ...base, pts: g.pts.map(p => [p.x, p.y]) };
  return { ...base, x: Math.min(s.x, c.x), y: Math.min(s.y, c.y), w: Math.abs(c.x - s.x), h: Math.abs(c.y - s.y) };
}
function commitText() {
  const a = state.annot, input = document.getElementById('ann-text'); if (!a?.textAt) return;
  const text = (input?.value || '').trim(), at = a.textAt; a.textAt = null;
  if (text) { annPush(); a.marks.push({ id: uid(), t: 'text', x: at.x, y: at.y, c: a.color, s: a.size, text }); }
  renderAnnot();
}
document.addEventListener('pointerdown', (e) => {
  const a = state.annot; if (!a || !e.target.closest?.('#ann-svg')) return;
  e.preventDefault();
  if (a.textAt) { commitText(); return; }
  const pt = annPt(e), t = a.tool;
  if (t === 'select') {
    const g = e.target.closest('[data-mark]'); a.sel = g ? g.dataset.mark : null;
    if (g) annG = { kind: 'move', start: pt, orig: JSON.parse(JSON.stringify(a.marks.find(m => m.id === a.sel))), moved: false };
    redrawMarks(); return;
  }
  if (t === 'text') { a.textAt = pt; renderAnnot(); return; }
  if (t === 'badge') { annPush(); a.marks.push({ id: uid(), t: 'badge', x: pt.x, y: pt.y, c: a.color, s: a.size, n: a.marks.filter(m => m.t === 'badge').length + 1 }); redrawMarks(); return; }
  annG = { kind: 'draw', t, start: pt, cur: pt, pts: [pt] };
});
document.addEventListener('pointermove', (e) => {
  const a = state.annot; if (!a || !annG) return;
  const pt = annPt(e), g = annG, H = annH();
  if (g.kind === 'move') {
    const m = a.marks.find(x => x.id === a.sel); if (!m) return;
    if (!g.moved && Math.hypot(pt.x - g.start.x, pt.y - g.start.y) < 0.004) return;
    if (!g.moved) { annPush(); g.moved = true; }
    Object.assign(m, JSON.parse(JSON.stringify(g.orig))); moveMark(m, pt.x - g.start.x, pt.y - g.start.y);
    const gm = document.getElementById('ann-marks'); if (gm) gm.innerHTML = a.marks.map(x => markSvg(x, H, true)).join(''); drawSel(); return;
  }
  g.cur = pt;
  if (g.t === 'pen') { const last = g.pts.at(-1); if (Math.hypot(pt.x - last.x, pt.y - last.y) > 0.003) g.pts.push(pt); }
  if (g.t === 'crop') { const s = g.start; a.crop = { x: Math.min(s.x, pt.x), y: Math.min(s.y, pt.y), w: Math.abs(pt.x - s.x), h: Math.abs(pt.y - s.y) }; renderCropOverlay(); return; }
  const d = document.getElementById('ann-draft'); if (d) d.innerHTML = markSvg(draftOf(g), H, false);
});
function renderCropOverlay() {
  const a = state.annot, svg = document.getElementById('ann-svg'), H = annH(); if (!svg) return;
  svg.querySelectorAll('.crop-mask,.crop-frame').forEach(x => x.remove());
  const c = a.crop; if (!c) return;
  svg.insertAdjacentHTML('beforeend', `<path class="crop-mask" fill-rule="evenodd" d="M0 0H1000V${H}H0Z M${c.x * 1000} ${c.y * H}h${c.w * 1000}v${c.h * H}h${-c.w * 1000}Z"/><rect class="crop-frame" x="${c.x * 1000}" y="${c.y * H}" width="${c.w * 1000}" height="${c.h * H}"/>`);
}
document.addEventListener('pointerup', () => {
  const a = state.annot, g = annG; annG = null; if (!a || !g) return;
  if (g.kind === 'move') { redrawMarks(); return; }
  const d = document.getElementById('ann-draft'); if (d) d.innerHTML = '';
  if (g.t === 'crop') { renderAnnot(); return; }
  const m = draftOf(g);
  const tiny = (m.t === 'arrow' && Math.hypot(m.x2 - m.x1, m.y2 - m.y1) < 0.012) || (m.t === 'pen' && m.pts.length < 2) || (('w' in m) && (m.w < 0.006 || m.h < 0.006));
  if (tiny) return;
  annPush(); a.marks.push(m); a.sel = null; redrawMarks();
});
document.addEventListener('keydown', (e) => {
  const a = state.annot; if (!a) return;
  if (e.target.id === 'ann-text') { if (e.key === 'Enter') { e.preventDefault(); commitText(); } else if (e.key === 'Escape') { a.textAt = null; renderAnnot(); } return; }
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); e.shiftKey ? annRedo() : annUndo(); }
  else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') { e.preventDefault(); annRedo(); }
  else if ((e.key === 'Delete' || e.key === 'Backspace') && a.sel) { e.preventDefault(); actions['ann-delete'](); }
  else if (e.key === 'Escape' && !a.dirty) actions['ann-cancel']();
});
document.addEventListener('wheel', (e) => {
  if (!state.annot || !(e.ctrlKey || e.metaKey) || !e.target.closest?.('#annot-scroll')) return;
  e.preventDefault(); annZoom(state.annot.zoom + (e.deltaY < 0 ? 0.2 : -0.2));
}, { passive: false });
function annZoom(z) { state.annot.zoom = Math.min(4, Math.max(0.5, Math.round(z * 100) / 100)); renderAnnot(); }

async function openAnnotate(id) {
  const p = currentProject(), b = curSlide()?.blocks.find(x => x.id === id);
  if (!p || !b || !isImg(b)) return;
  const srcId = b.srcAssetId || b.assetId;
  const blob = await getAssetBlob(p.id, srcId).catch(() => null);
  const bmp = blob ? await createImageBitmap(blob).catch(() => null) : null;
  if (!bmp) { notify('تعذّر فتح هذه الصورة للتعليم (النوع غير مدعوم).'); return; }
  state.annot = { blockId: id, srcAssetId: srcId, nw: bmp.width, nh: bmp.height, srcBlob: blob, srcUrl: URL.createObjectURL(blob), marks: JSON.parse(JSON.stringify(b.marks || [])), tool: 'arrow', color: '#e53935', size: 1, sel: null, zoom: 1, crop: null, cropped: false, dirty: false, confirmCancel: false, past: [], future: [], textAt: null, urls: [] };
  bmp.close?.(); render();
}
function closeAnnotate() { state.annot = null; annG = null; render(); }
async function saveAnnotate() {
  const a = state.annot, p = currentProject(), b = curSlide()?.blocks.find(x => x.id === a.blockId);
  if (!a || !p || !b) return closeAnnotate();
  if (!a.dirty) return closeAnnotate();
  const btn = document.getElementById('ann-save'); if (btn) btn.disabled = true;
  try {
    checkpoint();
    let srcId = a.srcAssetId;
    if (a.cropped) {
      srcId = uid(); await putAssetBlob(p.id, srcId, a.srcBlob);
      p.imageAssets[srcId] = { id: srcId, fileName: `${p.imageAssets[a.srcAssetId]?.fileName || 'صورة'} (مقصوصة)`, contentType: a.srcBlob.type, stored: true };
      state.assetUrls[srcId] = URL.createObjectURL(a.srcBlob);
    }
    if (!a.marks.length) { b.assetId = srcId; delete b.srcAssetId; delete b.marks; }
    else {
      const flat = await flattenMarks(a.srcBlob, a.marks, a.nw, a.nh), fid = uid();
      await putAssetBlob(p.id, fid, flat);
      p.imageAssets[fid] = { id: fid, fileName: `${p.imageAssets[srcId]?.fileName || 'صورة'} (معلَّمة)`, contentType: flat.type, stored: true };
      state.assetUrls[fid] = URL.createObjectURL(flat);
      b.assetId = fid; b.srcAssetId = srcId; b.marks = a.marks;
    }
    if (a.cropped && b.box) b.box.h = Math.min(96, b.box.w * (a.nh / a.nw) * STAGE_AR);
    touch(p); state.annot = null; state.notice = 'تم حفظ تعليم الصورة.'; render();
  } catch { state.annot.confirmCancel = false; state.error = 'تعذّر حفظ الصورة المعلَّمة.'; renderAnnot(); }
}
async function applyCrop() {
  const a = state.annot, c = a.crop; if (!c || c.w < 0.03 || c.h < 0.03) { a.crop = null; renderAnnot(); return; }
  annPush();
  const bmp = await createImageBitmap(a.srcBlob), sx = Math.round(c.x * bmp.width), sy = Math.round(c.y * bmp.height), sw = Math.max(1, Math.round(c.w * bmp.width)), sh = Math.max(1, Math.round(c.h * bmp.height));
  const cv = document.createElement('canvas'); cv.width = sw; cv.height = sh;
  const ctx = cv.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, sw, sh); ctx.drawImage(bmp, sx, sy, sw, sh, 0, 0, sw, sh); bmp.close?.();
  const type = a.srcBlob.type === 'image/png' ? 'image/png' : 'image/jpeg';
  const blob = await new Promise(r => cv.toBlob(r, type, 0.92));
  a.marks.forEach(m => remapMark(m, c));
  a.srcBlob = blob; a.srcUrl = URL.createObjectURL(blob); a.nw = sw; a.nh = sh; a.cropped = true; a.crop = null; a.tool = 'arrow'; a.sel = null;
  renderAnnot();
}

const annotActions = {
  'annotate': (t) => { void openAnnotate(t.dataset.id); },
  'ann-tool': (t) => { const a = state.annot; a.tool = t.dataset.tool; a.sel = null; a.textAt = null; if (a.tool !== 'crop') a.crop = null; renderAnnot(); },
  'ann-color': (t) => { const a = state.annot; a.color = t.dataset.color; const m = a.sel && a.marks.find(x => x.id === a.sel); if (m) { annPush(); m.c = a.color; } renderAnnot(); },
  'ann-size': (t) => { const a = state.annot; a.size = Number(t.dataset.size); const m = a.sel && a.marks.find(x => x.id === a.sel); if (m) { annPush(); m.s = a.size; } renderAnnot(); },
  'ann-undo': () => annUndo(), 'ann-redo': () => annRedo(),
  'ann-delete': () => { const a = state.annot; if (!a.sel) return; annPush(); a.marks = a.marks.filter(m => m.id !== a.sel); a.sel = null; redrawMarks(); },
  'ann-clear': () => { const a = state.annot; if (!a.marks.length) return; annPush(); a.marks = []; a.sel = null; redrawMarks(); },
  'ann-zoom-in': () => annZoom(state.annot.zoom + 0.25), 'ann-zoom-out': () => annZoom(state.annot.zoom - 0.25), 'ann-zoom-fit': () => annZoom(1),
  'ann-crop-apply': () => { void applyCrop(); },
  'ann-crop-cancel': () => { state.annot.crop = null; renderAnnot(); },
  'ann-save': () => { void saveAnnotate(); },
  'ann-cancel': () => { const a = state.annot; if (a.dirty && !a.confirmCancel) { a.confirmCancel = true; renderAnnot(); return; } closeAnnotate(); },
};
/* ================= المزامنة مع Supabase (اختيارية) =================
 * التخزين المحلي يبقى الأساس؛ المزامنة تنسخ المشاريع والصور إلى حسابك وتجلبها على أي جهاز.
 * تستخدم واجهات Supabase REST مباشرة (بلا مكتبات)، والأمان عبر RLS في قاعدة البيانات.
 */
const SB_BUCKET = 'waraq-assets';
const SB_CONFIG = { url: 'https://jwfylfkwjnznvnvldeiu.supabase.co', key: 'sb_publishable_aAIRikU5ZI0ki9j5Ls8Y6g_kZ95pzBL' }; // المفتاح العام (publishable) آمن للعرض؛ الحماية في سياسات RLS وتسجيل الدخول
const lsGet = (k, d) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch { return d; } };
const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* التخزين غير متاح */ } };
const sync = { cfg: SB_CONFIG, session: lsGet('waraq-sb-session', null), status: 'idle', last: lsGet('waraq-sb-last', ''), msg: '', open: false, busy: false, step: '' };
let syncTimer = 0, syncRunning = false, syncAgain = false;
const syncConfigured = () => !!(sync.cfg.url && sync.cfg.key);
const syncOn = () => syncConfigured() && !!sync.session;
const sbBase = () => sync.cfg.url.replace(/\/+$/, '');

async function sbAuthed() {
    const s = sync.session; if (!s) throw new Error('no-session');
    if (s.expires_at && s.expires_at * 1000 - Date.now() < 60000 && s.refresh_token) {
        const r = await fetch(`${sbBase()}/auth/v1/token?grant_type=refresh_token`, { method: 'POST', headers: { apikey: sync.cfg.key, 'Content-Type': 'application/json' }, body: JSON.stringify({ refresh_token: s.refresh_token }) });
        if (!r.ok) { forceLogout(false); throw new Error('session-expired'); }
        sync.session = sbSessionOf(await r.json()); lsSet('waraq-sb-session', sync.session);
    }
    return sync.session.access_token;
}
const sbSessionOf = (j) => ({ access_token: j.access_token, refresh_token: j.refresh_token, expires_at: j.expires_at || Math.floor(Date.now() / 1000) + (j.expires_in || 3600), user: { id: j.user?.id, email: j.user?.email } });
async function sb(path, { method = 'GET', body, headers = {}, raw = false } = {}) {
    const token = await sbAuthed();
    const r = await fetch(`${sbBase()}${path}`, { method, headers: { apikey: sync.cfg.key, Authorization: `Bearer ${token}`, ...(body && !(body instanceof Blob) ? { 'Content-Type': 'application/json' } : {}), ...headers }, body: body instanceof Blob ? body : body ? JSON.stringify(body) : undefined });
    if (!r.ok) { const e = new Error(`HTTP ${r.status}`); e.status = r.status; e.body = await r.text().catch(() => ''); throw e; }
    return raw ? r : (r.status === 204 ? null : r.json().catch(() => null));
}
const msOf = (v) => { const t = Date.parse(v || ''); return Number.isFinite(t) ? t : 0; };
const tombs = () => lsGet('waraq-sb-tombs', []);
const addTombstone = (id) => { const t = tombs(); if (!t.includes(id)) { t.push(id); lsSet('waraq-sb-tombs', t); } };
const uploaded = () => new Set(lsGet('waraq-sb-uploaded', []));
const assetIdsOf = (p) => { const s = new Set(); for (const sl of p.slides || []) for (const b of sl.blocks || []) { if (b.assetId) s.add(b.assetId); if (b.srcAssetId) s.add(b.srcAssetId); } return [...s]; };
const assetPath = (pid, aid) => `${sync.session.user.id}/${encodeURIComponent(pid)}/${encodeURIComponent(aid)}`;

async function pushProject(p) {
    const up = uploaded(), uid0 = sync.session.user.id;
    for (const aid of assetIdsOf(p)) {
        const key = `${uid0}/${p.id}/${aid}`; if (up.has(key)) continue;
        const blob = await getAssetBlob(p.id, aid); if (!blob) continue;
        await sb(`/storage/v1/object/${SB_BUCKET}/${assetPath(p.id, aid)}`, { method: 'POST', body: blob, headers: { 'Content-Type': blob.type || 'application/octet-stream', 'x-upsert': 'true' } });
        up.add(key); lsSet('waraq-sb-uploaded', [...up]);
    }
    const data = { ...p }; delete data._rev;
    await sb('/rest/v1/waraq_projects?on_conflict=user_id,id', { method: 'POST', headers: { Prefer: 'resolution=merge-duplicates,return=minimal' }, body: { id: p.id, data, updated_at: p.updatedAt, deleted: false } });
}
async function pullProject(id) {
    const rows = await sb(`/rest/v1/waraq_projects?id=eq.${encodeURIComponent(id)}&select=data,deleted`);
    const row = rows?.[0]; if (!row || row.deleted) return null;
    const p = row.data; await saveProject(p, true);
    const up = uploaded();
    for (const aid of assetIdsOf(p)) {
        if (await getAssetBlob(p.id, aid)) continue;
        try {
            const r = await sb(`/storage/v1/object/authenticated/${SB_BUCKET}/${assetPath(p.id, aid)}`, { raw: true });
            await putAssetBlob(p.id, aid, await r.blob()); up.add(`${sync.session.user.id}/${p.id}/${aid}`);
        } catch { /* صورة غير موجودة على الخادم */ }
    }
    lsSet('waraq-sb-uploaded', [...up]);
    return p;
}
async function pushTombstone(id) {
    await sb('/rest/v1/waraq_projects?on_conflict=user_id,id', { method: 'POST', headers: { Prefer: 'resolution=merge-duplicates,return=minimal' }, body: { id, data: {}, updated_at: new Date().toISOString(), deleted: true } });
    try {
        const list = await sb(`/storage/v1/object/list/${SB_BUCKET}`, { method: 'POST', body: { prefix: `${sync.session.user.id}/${encodeURIComponent(id)}`, limit: 1000 } });
        const names = (list || []).map(o => `${sync.session.user.id}/${encodeURIComponent(id)}/${o.name}`);
        if (names.length) await sb(`/storage/v1/object/${SB_BUCKET}`, { method: 'DELETE', body: { prefixes: names } });
    } catch { /* تنظيف الصور اختياري */ }
}
async function syncNow() {
    if (!syncOn()) return;
    if (syncRunning) { syncAgain = true; return; }
    syncRunning = true; sync.status = 'syncing'; sync.msg = ''; paintSync();
    try {
        await flushSave();
        for (const id of [...tombs()]) { await pushTombstone(id); lsSet('waraq-sb-tombs', tombs().filter(x => x !== id)); }
        const remote = await sb('/rest/v1/waraq_projects?select=id,updated_at,deleted') || [];
        const rmap = new Map(remote.map(r => [r.id, r]));
        const locals = await getProjects(), lmap = new Map(locals.map(p => [p.id, p]));
        const pulled = new Set(), removed = new Set(); let changed = 0;
        for (const r of remote) {
            const l = lmap.get(r.id), rt = msOf(r.updated_at), lt = l ? msOf(l.updatedAt) : 0;
            if (r.deleted) { if (l && lt <= rt) { await deleteProject(r.id); removed.add(r.id); changed++; } else if (l) { await pushProject(l); changed++; } continue; }
            if (!l || rt > lt) { if (await pullProject(r.id)) { pulled.add(r.id); changed++; } }
            else if (rt < lt) { await pushProject(l); changed++; }
        }
        for (const l of locals) if (!rmap.has(l.id)) { await pushProject(l); changed++; }
        if (pulled.size || removed.size) {
            const keepId = state.projectId, keepMode = state.mode;
            state.projects = await getProjects(); for (const id of [...pulled, ...removed]) analysisCache.delete(id);
            if (removed.has(keepId)) { state.projectId = ''; releaseAssetUrls(); if (state.projects[0]) await openProject(state.projects[0].id); }
            else if (pulled.has(keepId) && !state.annot) await openProject(keepId, keepMode);
        }
        sync.status = 'ok'; sync.last = new Date().toISOString(); lsSet('waraq-sb-last', sync.last);
        sync.msg = changed ? `تمت المزامنة (${changed} تغيير).` : 'كل شيء محدّث.';
    } catch (e) {
        sync.status = 'error';
        sync.msg = e.message === 'session-expired' || e.status === 401 ? 'انتهت الجلسة؛ سجّل الدخول من جديد.' : e.status === 404 || /relation|schema cache/.test(e.body || '') ? 'لم يُنفَّذ إعداد قاعدة البيانات بعد (SQL).' : e.name === 'TypeError' ? 'تعذّر الاتصال بالإنترنت أو بعنوان المشروع.' : `تعذّرت المزامنة (${e.status || e.message}).`;
    }
    syncRunning = false;
    if (sync.status === 'error' && /انتهت الجلسة/.test(sync.msg)) { forceLogout(false); return; }
    render();
    if (syncAgain) { syncAgain = false; queueSync(500); }
}
function queueSync(ms = 4000) { if (!syncOn()) return; clearTimeout(syncTimer); syncTimer = window.setTimeout(() => { syncTimer = 0; void syncNow(); }, ms); }
function paintSync() { const b = document.getElementById('sync-btn'); if (b) b.outerHTML = syncButtonHtml(); }
function syncButtonHtml() {
    const st = sync.status === 'syncing' ? 'busy' : sync.status === 'error' ? 'err' : 'ok';
    const tip = st === 'err' ? (sync.msg || 'تعذّرت المزامنة') : st === 'busy' ? 'تتم المزامنة…' : 'محفوظ';
    return `<button class="sync-btn sync-${st}" id="sync-btn" data-action="sync-signout" title="${h(tip)} — انقر لتسجيل الخروج"><span class="sync-dot"></span><span>خروج</span></button>`;
}
const syncModalHtml = () => '';
/** تسجيل الخروج: مزامنة أخيرة ثم مسح البيانات المحلية من هذا المتصفح. */
function forceLogout(wipe = true) { lsSet('waraq-sb-session', null); if (wipe) lsSet('waraq-wipe', 1); location.reload(); }
const syncActions = {
    'sync-signout': async () => {
        await syncNow();
        if (sync.status === 'error' && !window.confirm('تعذّرت المزامنة الأخيرة. الخروج سيمسح من هذا المتصفح أي تعديلات لم تُرفع بعد. هل تريد المتابعة؟')) { render(); return; }
        forceLogout();
    },
};
window.addEventListener('online', () => queueSync(500));
document.addEventListener('visibilitychange', () => { if (!document.hidden) queueSync(800); });

/* ===== بوابة الدخول: لا يظهر شيء من النظام قبل تسجيل الدخول ===== */
async function wipeLocalData() {
    for (const k of Object.keys(localStorage)) if (k.startsWith('waraq-')) localStorage.removeItem(k);
    try { sessionStorage.clear(); } catch { /* غير متاح */ }
    await new Promise(res => { try { const r = indexedDB.deleteDatabase('waraq-analysis'); r.onsuccess = r.onerror = r.onblocked = () => res(); } catch { res(); } });
}
function gateHtml() {
    return `<main class="gate"><form id="gate-form" class="gate-card" autocomplete="on"><h1>تسجيل الدخول</h1>
        <label for="gate-email">البريد الإلكتروني</label><input id="gate-email" type="email" autocomplete="username" required dir="ltr">
        <label for="gate-pass">كلمة المرور</label><input id="gate-pass" type="password" autocomplete="current-password" required dir="ltr">
        <div class="gate-err" id="gate-err" role="alert"></div><button type="submit" id="gate-go">دخول</button></form></main>`;
}
document.addEventListener('submit', async (ev) => {
    if (ev.target.id !== 'gate-form') return;
    ev.preventDefault();
    const err = document.getElementById('gate-err'), go = document.getElementById('gate-go'); err.textContent = ''; go.disabled = true;
    try {
        const r = await fetch(`${sbBase()}/auth/v1/token?grant_type=password`, { method: 'POST', headers: { apikey: sync.cfg.key, 'Content-Type': 'application/json' }, body: JSON.stringify({ email: document.getElementById('gate-email').value.trim(), password: document.getElementById('gate-pass').value }) });
        if (!r.ok) throw new Error('bad');
        lsSet('waraq-sb-session', sbSessionOf(await r.json())); location.reload();
    } catch (e) { err.textContent = e.message === 'bad' ? 'بيانات الدخول غير صحيحة.' : 'تعذّر الاتصال. حاول لاحقًا.'; go.disabled = false; }
});

/* ================= السلوك ================= */
function resetFilters() { state.ownerFilter = 'الكل'; state.severity = 'الكل'; state.sectionFilter = 'الكل'; state.statusFilter = 'الكل'; state.slideFilter = ''; state.search = ''; }
function snapshotIfChanged(p) {
  const s = analysisOf(p).summary;
  const sig = `${s.cases.length}|${s.bySeverity['مرتفع']}|${s.byStatus.done}|${p.updatedAt}`;
  const last = (p.history || []).at(-1);
  if (last && last.sig === sig) return;
  p.history = [...(p.history || []), { at: new Date().toISOString(), sig, total: s.cases.length, high: s.bySeverity['مرتفع'], done: s.byStatus.done, progress: s.progress, risk: s.risk, slides: p.slides.length }].slice(-30);
  scheduleSave(p);
}
async function openProject(id, mode) {
  await flushSave();
  const p = state.projects.find(x => x.id === id); if (!p) return;
  state.view = 'project'; state.projectId = id; state.slideIdx = 0; state.sel = null; state.editingId = null; state.zoom = 1; state.hist = { past: [], future: [] }; state.annot = null; state.ownerFilter = 'الكل'; state.present = null; state.sidebarOpen = false;
  state.selectedFinding = null; state.selectedSlide = null; state.lightbox = null; state.undo = null;
  state.tab = 'cases'; resetFilters(); state.slidesFilter = 'all';
  state.mode = mode || (p.slides.some(s => s.blocks.length) ? 'analyze' : 'edit');
  if (state.mode === 'analyze') snapshotIfChanged(p);
  releaseAssetUrls(); render();
  await loadAssetUrls(id);
}
function setMode(mode) {
  const p = currentProject(); if (!p) return;
  state.mode = mode; state.sel = null; state.editingId = null;
  if (mode === 'analyze') snapshotIfChanged(p);
  render();
  window.scrollTo(0, 0);
}
async function importFiles(files) {
  const selected = Array.from(files);
  const valid = selected.filter(f => /\.(pptx|json)$/i.test(f.name));
  if (!valid.length) { state.error = 'اختر ملف PowerPoint (.pptx) أو نسخة احتياطية (.json).'; render(); return; }
  state.error = selected.length > valid.length ? 'تم تجاهل الملفات غير المدعومة.' : '';
  let lastId = '';
  for (let i = 0; i < valid.length; i++) {
    const file = valid[i];
    state.progress = { name: file.name, index: i + 1, total: valid.length }; render();
    try {
      let project;
      if (/\.json$/i.test(file.name)) project = await backupToProject(JSON.parse(await file.text()));
      else {
        const { slides, imageAssets, imageStats } = await readPowerPoint(file, { allImages: state.allImages });
        project = slidesToProject({ name: file.name.replace(/\.pptx$/i, ''), slides, imageAssets, imageStats });
        for (const id of Object.keys(project.imageAssets)) if (imageAssets[id]?.blob) await putAssetBlob(project.id, id, imageAssets[id].blob);
      }
      await saveProject(project);
      state.projects = [project, ...state.projects]; lastId = project.id;
    } catch (reason) {
      state.error = reason instanceof SyntaxError ? `${file.name}: الملف ليس JSON صالحًا.` : reason instanceof Error ? `${file.name}: ${reason.message}` : `تعذّر استيراد ${file.name}.`;
    }
  }
  state.progress = null;
  if (lastId) { state.notice = 'تم استيراد المشروع وتحليله.'; await openProject(lastId, 'analyze'); } else render();
}
async function duplicateProject(id) {
  const src = state.projects.find(p => p.id === id); if (!src) return;
  await flushSave();
  const copy = { ...JSON.parse(JSON.stringify(src)), id: uid(), name: `${src.name} (نسخة)`, history: [], createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() };
  delete copy._rev;
  try { await copyAssetBlobs(src.id, copy.id); await saveProject(copy); state.projects = [copy, ...state.projects]; state.notice = 'تم نسخ المشروع.'; await openProject(copy.id, 'edit'); }
  catch { state.error = 'تعذّر نسخ المشروع.'; render(); }
}
async function confirmDeleteProject() {
  const target = state.deleteProject; if (!target) return;
  try {
    await deleteProject(target.id);
    addTombstone(target.id); queueSync(1000);
    analysisCache.delete(target.id);
    state.projects = state.projects.filter(p => p.id !== target.id);
    state.deleteProject = null; state.notice = 'تم حذف المشروع وصوره من هذا المتصفح.';
    if (state.projectId === target.id) {
      if (state.projects[0]) { await openProject(state.projects[0].id); return; }
      state.projectId = ''; releaseAssetUrls();
    }
  } catch { state.notice = 'تعذّر حذف المشروع.'; state.deleteProject = null; }
  render();
}
async function saveProjectForm() {
  const d = state.projectForm; if (!d || !d.name.trim()) return;
  if (d.mode === 'new') {
    const project = newProject(d);
    try { await saveProject(project); } catch { state.error = 'تعذّر حفظ المشروع في هذا المتصفح.'; render(); return; }
    state.projects = [project, ...state.projects]; state.projectForm = null;
    await openProject(project.id, 'edit');
    state.focusId = 'slide-title'; render();
  } else {
    const p = state.projects.find(x => x.id === d.id); if (!p) return;
    Object.assign(p, { name: d.name.trim(), owner: d.owner.trim(), location: d.location.trim(), description: d.description.trim() });
    touch(p); state.projectForm = null; state.notice = 'تم تحديث معلومات المشروع.'; render();
  }
}
function jumpToEdit(slideNumber, blockId) {
  state.selectedFinding = null; state.selectedSlide = null;
  state.mode = 'edit'; state.slideIdx = Math.max(0, slideNumber - 1); state.sel = blockId || null; state.editingId = null;
  render();
  const target = blockId && document.getElementById(`block-${blockId}`);
  if (target) { target.scrollIntoView({ block: 'center' }); target.classList.add('flash'); }
  else window.scrollTo(0, 0);
}
const findCase = (id) => activeSummary()?.cases.find(c => c.id === id);
const lightboxStep = (delta) => { const a = active(); const n = a ? galleryItems(a).length : 0; if (n) { state.lightbox = (state.lightbox + delta + n) % n; render(); } };
const presentStep = (d) => { const p = currentProject(); if (p && state.present !== null) { state.present = Math.max(0, Math.min(p.slides.length - 1, state.present + d)); render(); } };
const notify = (text) => { state.notice = text; render(); };

const actions = {
  ...editorActions, ...stageActions, ...annotActions, ...syncActions,
  'open-board': () => { flushSave().then(() => { state.view = 'board'; state.sidebarOpen = false; render(); window.scrollTo(0, 0); }); },
  'board-owner': (t) => { state.boardOwner = t.dataset.owner; render(); },
  'board-open': async (t) => { await openProject(t.dataset.project, 'edit'); jumpToEdit(Number(t.dataset.slide), t.dataset.block); },
  'filter-owner': (t) => { state.ownerFilter = state.ownerFilter === t.dataset.owner ? 'الكل' : t.dataset.owner; render(); },
  'toggle-side-collapse': () => { state.sideCollapsed = !state.sideCollapsed; render(); },
  'toggle-sidebar': () => { state.sidebarOpen = !state.sidebarOpen; render(); },
  'new-project': () => { state.sidebarOpen = false; state.projectForm = { mode: 'new', name: '', owner: '', location: '', description: '', template: 'blank' }; state.focusId = 'pf-name'; render(); },
  'edit-info': () => { const p = currentProject(); if (!p) return; state.projectForm = { mode: 'edit', id: p.id, name: p.name, owner: p.owner || '', location: p.location || '', description: p.description || '' }; state.focusId = 'pf-name'; render(); },
  'close-form': () => { state.projectForm = null; render(); },
  'save-form': () => saveProjectForm(),
  'import': () => $fileInput.click(),
  'open-project': (t) => openProject(t.dataset.id, t.dataset.mode),
  'dup-project': (t) => duplicateProject(t.dataset.id),
  'ask-delete-project': (t) => { state.deleteProject = state.projects.find(p => p.id === t.dataset.id) || null; render(); },
  'confirm-delete-project': () => confirmDeleteProject(),
  'mode-edit': () => setMode('edit'),
  'mode-analyze': async () => { await flushSave(); setMode('analyze'); },
  'save-analyze': async () => { await flushSave(); setMode('analyze'); notify('تم حفظ المشروع وتحليله.'); },
  'mode-present': () => { const p = currentProject(); if (p) { state.present = Math.min(state.slideIdx, p.slides.length - 1); render(); } },
  'present-close': () => { state.present = null; render(); },
  'present-next': () => presentStep(1),
  'present-prev': () => presentStep(-1),
  'export-csv': () => { const a = active(); if (a) { downloadCasesCsv(a); notify('تم تجهيز ملف CSV للحالات.'); } },
  'export-backup': async () => { const p = currentProject(); if (!p) return; try { await downloadBackup(p); notify('تم تنزيل النسخة الاحتياطية (مع الصور).'); } catch { notify('تعذّر تجهيز النسخة الاحتياطية.'); } },
  'export-pptx': async () => { const p = currentProject(); if (!p) return; if (ensureAllLayouts(p)) scheduleSave(p); await flushSave(); try { await downloadPptx(p); notify('تم تصدير المشروع إلى PowerPoint (.pptx).'); } catch (e) { console.error(e); notify('تعذّر تصدير ملف PowerPoint.'); } },
  'export-html': async () => { const a = active(); if (!a) return; try { await downloadAnalysisReport(a); notify('تم تنزيل التقرير المستقل؛ يتضمن الملخص والحالات والشرائح والصور.'); } catch { notify('تعذّر تجهيز التقرير. حاول مرة أخرى.'); } },
  'dismiss-error': () => { state.error = ''; render(); },
  'dismiss-notice': () => { state.notice = ''; render(); },
  'chart-severity': (t) => { state.severity = state.severity === t.dataset.value ? 'الكل' : t.dataset.value; state.tab = 'cases'; render(); },
  'chart-section': (t) => { state.sectionFilter = state.sectionFilter === t.dataset.value ? 'الكل' : t.dataset.value; state.tab = 'cases'; render(); },
  'goto-slide': (t) => { resetFilters(); state.slideFilter = t.dataset.id; state.tab = 'cases'; render(); document.getElementById('workspace')?.scrollIntoView({ behavior: 'smooth', block: 'start' }); },
  'goto-case': (t) => { const c = findCase(t.dataset.id); if (!c) return; resetFilters(); state.slideFilter = String(c.slideNumber); state.tab = 'cases'; render(); document.getElementById('workspace')?.scrollIntoView({ behavior: 'smooth', block: 'start' }); },
  'tab': (t) => { state.tab = t.dataset.value; render(); },
  'clear-filters': () => { resetFilters(); render(); },
  'open-case': (t) => { const c = findCase(t.dataset.id); if (!c) return; state.selectedFinding = c.observation ?? c.findings[0]; state.selectedSlide = null; render(); },
  'edit-case': (t) => { const c = findCase(t.dataset.id); if (c) jumpToEdit(c.slideNumber, c.blockId); },
  'goto-comment': (t) => jumpToEdit(Number(t.dataset.slide), t.dataset.id),
  'open-slide': (t) => { state.selectedSlide = active()?.slides.find(x => x.number === Number(t.dataset.id)) || null; state.selectedFinding = null; render(); },
  'close-quote': () => { state.selectedFinding = null; state.selectedSlide = null; render(); },
  'edit-from-quote': () => { const f = state.selectedFinding, sl = state.selectedSlide; jumpToEdit(f?.slideNumber || sl?.number || 1, f?.blockId); },
  'lightbox': (t) => { const i = Number(t.dataset.index); if (i >= 0) { state.lightbox = i; render(); } },
  'close-lightbox': () => { state.lightbox = null; render(); },
  'lb-next': () => lightboxStep(1),
  'lb-prev': () => lightboxStep(-1),
  'cancel-delete': () => { state.deleteSlide = null; state.deleteProject = null; render(); },
};

document.addEventListener('click', (event) => {
  const backdrop = event.target.closest?.('[data-backdrop]');
  if (backdrop && event.target === backdrop) { actions[backdrop.dataset.backdrop]?.(backdrop); return; }
  const trigger = event.target.closest?.('[data-action]');
  if (!trigger || trigger.disabled) return;
  actions[trigger.dataset.action]?.(trigger, event);
});
document.addEventListener('keydown', (event) => {
  if (state.annot) return;
  if (stageKey(event)) return;
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
    event.preventDefault();
    if (currentProject()) void flushSave().then(() => notify('تم حفظ المشروع.'));
    return;
  }
  if (state.present !== null) {
    if (event.key === 'Escape') actions['present-close']();
    else if (event.key === 'ArrowLeft' || event.key === ' ' || event.key === 'PageDown') { event.preventDefault(); presentStep(1); }
    else if (event.key === 'ArrowRight' || event.key === 'PageUp') presentStep(-1);
    return;
  }
  if (event.key === 'Escape') {
    if (state.deleteSlide || state.deleteProject) actions['cancel-delete']();
    else if (state.projectForm) actions['close-form']();
    else if (state.lightbox !== null) actions['close-lightbox']();
    else if (state.selectedFinding || state.selectedSlide) actions['close-quote']();
    else if (state.sidebarOpen) actions['toggle-sidebar']();
  } else if (state.lightbox !== null && event.key === 'ArrowRight') lightboxStep(-1);
  else if (state.lightbox !== null && event.key === 'ArrowLeft') lightboxStep(1);
  else if (event.key === 'Enter' && state.projectForm && event.target.id?.startsWith('pf-') && event.target.tagName === 'INPUT') { event.preventDefault(); void saveProjectForm(); }
});
document.addEventListener('input', (event) => {
  const t = event.target, id = t.id;
  if (id === 'search') { state.search = t.value; renderCasesOnly(); }
  else if (id === 'slide-filter') { state.slideFilter = t.value; renderCasesOnly(); }
  else if (id === 'project-search') { state.projectSearch = t.value; const l = document.getElementById('project-list'); if (l) l.innerHTML = projectListHtml(); const a2 = document.getElementById('analysis-list'); if (a2) a2.innerHTML = analysisListHtml(); }
  else if (id?.startsWith('pf-') && state.projectForm) {
    state.projectForm[id.slice(3)] = t.value;
    if (id === 'pf-name') { const b = document.getElementById('save-form'); if (b) b.disabled = !t.value.trim(); }
  } else if (t.tagName !== 'SELECT' && (t.dataset.block || t.dataset.slideTitle || t.dataset.slideNotes)) {
    applyField(t);
    if (t.tagName === 'TEXTAREA') { t.style.height = 'auto'; t.style.height = `${t.scrollHeight + 2}px`; }
  }
});
document.addEventListener('change', (event) => {
  const t = event.target, id = t.id;
  if (id === 'severity') { state.severity = t.value; render(); }
  else if (id === 'section') { state.sectionFilter = t.value; render(); }
  else if (id === 'status-filter') { state.statusFilter = t.value; render(); }
  else if (id === 'sort-by') { state.sortBy = t.value; render(); }
  else if (id === 'group-by') { state.groupBy = t.value; render(); }
  else if (id === 'board-scope') { state.boardScope = t.value; render(); }
  else if (id === 'board-owner') { state.boardOwner = t.value; render(); }
  else if (id === 'board-project') { state.boardProject = t.value; render(); }
  else if (id === 'owner-filter') { state.ownerFilter = t.value; render(); }
  else if (t.classList?.contains('board-status')) {
    const p = state.projects.find(x => x.id === t.dataset.project); const c = p && analysisOf(p).summary.cases.find(x => x.id === t.dataset.case); if (!p || !c) return;
    if (c.manual) { for (const s of p.slides) { const b = s.blocks.find(x => x.id === c.blockId); if (b) b.status = t.value; } }
    else p.statuses = { ...(p.statuses || {}), [c.id]: t.value };
    p.updatedAt = new Date().toISOString(); p._rev = (p._rev || 0) + 1; void saveProject(p); render();
  }
  else if (id === 'slides-filter') { state.slidesFilter = t.value; render(); }
  else if (id === 'pf-template' && state.projectForm) state.projectForm.template = t.value;
  else if (id === 'opt-all-images') { state.allImages = t.checked; writeOption(t.checked); notify(t.checked ? 'سيُستورد كل ما في ملفات PowerPoint من صور.' : 'سيُستورد الأساسي من الصور فقط.'); }
  else if (t.classList?.contains('status-pick')) {
    const p = currentProject(); const c = findCase(t.dataset.case); if (!p || !c) return;
    if (c.manual) { for (const s of p.slides) { const b = s.blocks.find(x => x.id === c.blockId); if (b) b.status = t.value; } }
    else p.statuses = { ...(p.statuses || {}), [c.id]: t.value };
    touch(p); render();
  } else if (t.tagName === 'SELECT' && t.dataset.block) { if (applyField(t)) render(); }
  else if (t.type === 'date' && t.dataset.block) applyField(t);
});
document.addEventListener('error', (event) => {
  const img = event.target;
  if (img?.tagName === 'IMG' && img.closest('.slide-image-link')) img.closest('.slide-image-link').outerHTML = '<div class="slide-image-unavailable" role="img"><span>تعذّر عرض الصورة المضمّنة.</span></div>';
}, true);

$fileInput.addEventListener('change', () => { if ($fileInput.files) void importFiles($fileInput.files); $fileInput.value = ''; });
$imageInput.addEventListener('change', () => { const f = Array.from($imageInput.files || []); $imageInput.value = ''; if (f.length) void addImageFiles(f); });
document.addEventListener('paste', (event) => {
  if (state.mode !== 'edit' || state.projectForm) return;
  const files = Array.from(event.clipboardData?.files || []).filter(f => f.type.startsWith('image/'));
  if (files.length) { event.preventDefault(); void addImageFiles(files); }
});
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', (e) => {
  e.preventDefault();
  const files = Array.from(e.dataTransfer?.files || []);
  if (files.some(f => /\.(pptx|json)$/i.test(f.name))) void importFiles(files);
  else if (files.length && state.mode === 'edit') void addImageFiles(files);
});
window.addEventListener('beforeunload', () => { if (saveTimer) { const p = currentProject(); if (p) void saveProject(p); } });
document.addEventListener('visibilitychange', () => { if (document.hidden) void flushSave(); });

/* ================= البدء ================= */
(async function start() {
  if (lsGet('waraq-wipe', 0)) await wipeLocalData();
  if (!lsGet('waraq-sb-session', null)) { sync.session = null; render(); return; }
  try {
    const have = new Set((await getProjects()).map(p => p.id));
    for (const legacy of await getAnalyses().catch(() => [])) {
      try { if (!have.has(legacy.id)) await saveProject(await analysisToProject(legacy)); await removeLegacyRecord(legacy.id); } catch { /* يبقى التحليل القديم كما هو */ }
    }
    state.projects = await getProjects();
    state.loadingSaved = false;
    if (state.projects[0]) await openProject(state.projects[0].id); else render();
    queueSync(300);
  } catch { state.error = 'تعذّر فتح مساحة التخزين المحلية في هذا المتصفح.'; state.loadingSaved = false; render(); }
})();
render();

})();
