import type { ExtractedSlide } from "./types";

const POWERPOINT_NS =
  "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWING_NS =
  "http://schemas.openxmlformats.org/drawingml/2006/main";
const RELATIONSHIP_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_RELATIONSHIP_NS =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const MAX_FILE_BYTES = 150 * 1024 * 1024;
const MAX_SLIDE_XML_BYTES = 25 * 1024 * 1024;

interface ZipEntry {
  name: string;
  compressionMethod: number;
  compressedSize: number;
  uncompressedSize: number;
  localHeaderOffset: number;
}

function getZipEntries(bytes: Uint8Array): ZipEntry[] {
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
  if (
    entryCount === 0xffff ||
    centralDirectoryOffset === 0xffffffff ||
    entryCount > 25_000
  ) {
    throw new Error("هذا الملف يستخدم تنسيق ضغط غير مدعوم أو يحتوي على عناصر كثيرة جدًا.");
  }

  const decoder = new TextDecoder("utf-8");
  const entries: ZipEntry[] = [];
  let offset = centralDirectoryOffset;

  for (let index = 0; index < entryCount; index += 1) {
    if (
      offset + 46 > bytes.byteLength ||
      view.getUint32(offset, true) !== 0x02014b50
    ) {
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

async function readZipEntry(
  bytes: Uint8Array,
  entry: ZipEntry,
): Promise<Uint8Array> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const localOffset = entry.localHeaderOffset;

  if (
    localOffset + 30 > bytes.byteLength ||
    view.getUint32(localOffset, true) !== 0x04034b50
  ) {
    throw new Error(`تعذر فتح الملف الداخلي: ${entry.name}`);
  }

  if (entry.uncompressedSize > MAX_SLIDE_XML_BYTES) {
    throw new Error("تحتوي إحدى الشرائح على بيانات نصية أكبر من الحد المسموح.");
  }

  const nameLength = view.getUint16(localOffset + 26, true);
  const extraLength = view.getUint16(localOffset + 28, true);
  const dataStart = localOffset + 30 + nameLength + extraLength;
  const dataEnd = dataStart + entry.compressedSize;
  if (dataEnd > bytes.byteLength) {
    throw new Error(`بيانات الشريحة غير مكتملة: ${entry.name}`);
  }

  const compressed = bytes.subarray(dataStart, dataEnd);
  if (entry.compressionMethod === 0) return compressed.slice();
  if (entry.compressionMethod !== 8) {
    throw new Error("طريقة ضغط هذا الملف غير مدعومة.");
  }
  if (typeof DecompressionStream === "undefined") {
    throw new Error("المتصفح لا يدعم فك ضغط ملفات PowerPoint الحديثة.");
  }

  try {
    const input = new Blob([compressed.slice().buffer as ArrayBuffer]).stream();
    const decompressed = input.pipeThrough(
      new DecompressionStream("deflate-raw"),
    );
    return new Uint8Array(await new Response(decompressed).arrayBuffer());
  } catch {
    throw new Error(`تعذر فك ضغط بيانات الشريحة: ${entry.name}`);
  }
}

function parseXml(xml: string, label: string): Document {
  const document = new DOMParser().parseFromString(xml, "application/xml");
  if (document.getElementsByTagName("parsererror").length > 0) {
    throw new Error(`تعذر تحليل ملف ${label}.`);
  }
  return document;
}

function getOrderedSlidePaths(
  fileNames: Set<string>,
  presentationXml: string | undefined,
  relationshipsXml: string | undefined,
): string[] {
  const fallback = Array.from(fileNames).sort((a, b) => {
    const aNumber = Number(a.match(/slide(\d+)\.xml$/)?.[1] ?? 0);
    const bNumber = Number(b.match(/slide(\d+)\.xml$/)?.[1] ?? 0);
    return aNumber - bNumber;
  });

  if (!presentationXml || !relationshipsXml) return fallback;

  try {
    const presentation = parseXml(presentationXml, "ترتيب الشرائح");
    const relationships = parseXml(relationshipsXml, "روابط الشرائح");
    const targets = new Map<string, string>();

    for (const relationship of Array.from(
      relationships.getElementsByTagNameNS(
        PACKAGE_RELATIONSHIP_NS,
        "Relationship",
      ),
    )) {
      const id = relationship.getAttribute("Id");
      const target = relationship.getAttribute("Target");
      if (!id || !target) continue;
      const normalizedTarget = target.startsWith("/")
        ? target.slice(1)
        : target.startsWith("ppt/")
          ? target
          : `ppt/${target}`;
      targets.set(id, normalizedTarget);
    }

    const ordered = Array.from(
      presentation.getElementsByTagNameNS(POWERPOINT_NS, "sldId"),
    )
      .map((slide) => {
        const relationId = slide.getAttributeNS(RELATIONSHIP_NS, "id");
        return relationId ? targets.get(relationId) : undefined;
      })
      .filter((path): path is string => Boolean(path && fileNames.has(path)));

    return ordered.length > 0 ? ordered : fallback;
  } catch {
    return fallback;
  }
}

function extractText(xml: string, slideNumber: number): ExtractedSlide {
  const document = parseXml(xml, `الشريحة ${slideNumber}`);
  const shapes = Array.from(
    document.getElementsByTagNameNS(POWERPOINT_NS, "sp"),
  );
  const blocks = shapes
    .map((shape) =>
      Array.from(shape.getElementsByTagNameNS(DRAWING_NS, "p"))
        .map((paragraph) =>
          Array.from(paragraph.getElementsByTagNameNS(DRAWING_NS, "t"))
            .map((node) => node.textContent?.replace(/\s+/g, " ").trim() ?? "")
            .join(""),
        )
        .map((paragraph) => paragraph.trim())
        .filter(Boolean)
        .join("\n"),
    )
    .filter(Boolean);

  const text =
    blocks.length > 0
      ? blocks.join("\n")
      : Array.from(document.getElementsByTagNameNS(DRAWING_NS, "t"))
          .map((node) => node.textContent?.replace(/\s+/g, " ").trim() ?? "")
          .filter(Boolean)
          .join("\n");
  const title = blocks[0]?.slice(0, 120) ?? `الشريحة ${slideNumber}`;

  return { number: slideNumber, title, text };
}

export async function readPowerPoint(file: File): Promise<ExtractedSlide[]> {
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    throw new Error("يرجى اختيار ملف PowerPoint بامتداد ‎.pptx.");
  }
  if (file.size > MAX_FILE_BYTES) {
    throw new Error("حجم الملف أكبر من 150 ميغابايت.");
  }

  const bytes = new Uint8Array(await file.arrayBuffer());
  const entries = getZipEntries(bytes);
  const byName = new Map(entries.map((entry) => [entry.name, entry]));
  const slideNames = new Set(
    entries
      .filter((entry) => /^ppt\/slides\/slide\d+\.xml$/.test(entry.name))
      .map((entry) => entry.name),
  );

  if (slideNames.size === 0) {
    throw new Error("لم يتم العثور على شرائح قابلة للقراءة في هذا الملف.");
  }

  const readTextEntry = async (name: string) => {
    const entry = byName.get(name);
    if (!entry) return undefined;
    return new TextDecoder("utf-8").decode(await readZipEntry(bytes, entry));
  };

  const [presentationXml, relationshipsXml] = await Promise.all([
    readTextEntry("ppt/presentation.xml"),
    readTextEntry("ppt/_rels/presentation.xml.rels"),
  ]);
  const orderedPaths = getOrderedSlidePaths(
    slideNames,
    presentationXml,
    relationshipsXml,
  );
  const slides: ExtractedSlide[] = [];

  for (let index = 0; index < orderedPaths.length; index += 1) {
    const path = orderedPaths[index];
    const entry = byName.get(path);
    if (!entry) continue;
    const xml = new TextDecoder("utf-8").decode(
      await readZipEntry(bytes, entry),
    );
    slides.push(extractText(xml, index + 1));
  }

  return slides;
}
