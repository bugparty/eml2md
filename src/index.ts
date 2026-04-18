import { createHash } from "node:crypto";
import { mkdir, readFile, readdir, rename, writeFile } from "node:fs/promises";
import path from "node:path";

import * as cheerio from "cheerio";
import he from "he";
import { htmlToText } from "html-to-text";
import { simpleParser } from "mailparser";

export type DateValue = Date | string | null;

export type EmailAttachment = {
  filename: string;
  content: Buffer;
  contentType: string;
};

export type EmailPart = {
  date: DateValue;
  from: string;
  to: string;
  cc: string;
  subject: string;
  body: string;
  attachments: EmailAttachment[];
};

export type Logger = {
  info?: (message: string) => void;
  warn?: (message: string) => void;
  debug?: (message: string) => void;
  error?: (message: string) => void;
};

export type ConvertOptions = {
  newestFirst?: boolean;
  maxMarkdownChars?: number | null;
  logger?: Logger;
};

export type ConvertResult = {
  markdown: string;
  emails: EmailPart[];
};

type RequiredLogger = Required<Logger>;

type ParsedMailLike = {
  date?: Date | string | null;
  from?: { text?: string | null } | Array<{ text?: string | null }> | null;
  to?: { text?: string | null } | Array<{ text?: string | null }> | null;
  cc?: { text?: string | null } | Array<{ text?: string | null }> | null;
  subject?: string | null;
  text?: string | null;
  html?: string | false | null;
  attachments?: Array<{
    content?: Buffer | Uint8Array | string;
    contentType?: string;
    filename?: string;
  }>;
};

type BodyQualityMetrics = {
  chars: number;
  nonEmptyLines: number;
  informativeLines: number;
  noiseMarkers: number;
};

type CliOptions = {
  inputDir: string;
  outputDir: string;
  doneDir: string;
  newestFirst: boolean;
  verbose: boolean;
  quiet: boolean;
  keepInput: boolean;
  stdin: boolean;
  maxMarkdownChars?: number | null;
  help: boolean;
};

const CLOSE_LENGTH_RATIO_THRESHOLD = 0.05;

function noop(): void {}

function normalizeLogger(logger?: Logger): RequiredLogger {
  return {
    info: logger?.info ?? noop,
    warn: logger?.warn ?? noop,
    debug: logger?.debug ?? noop,
    error: logger?.error ?? noop,
  };
}

function parseArgs(argv: string[]): CliOptions {
  const options: CliOptions = {
    inputDir: "input",
    outputDir: "output",
    doneDir: "done",
    newestFirst: false,
    verbose: false,
    quiet: false,
    keepInput: false,
    stdin: false,
    maxMarkdownChars: null,
    help: false,
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];

    if (arg === "--input-dir" && argv[index + 1]) {
      options.inputDir = argv[index + 1];
      index += 1;
      continue;
    }
    if (arg === "--output-dir" && argv[index + 1]) {
      options.outputDir = argv[index + 1];
      index += 1;
      continue;
    }
    if (arg === "--done-dir" && argv[index + 1]) {
      options.doneDir = argv[index + 1];
      index += 1;
      continue;
    }
    if (arg === "--newest-first") {
      options.newestFirst = true;
      continue;
    }
    if (arg === "--keep-input") {
      options.keepInput = true;
      continue;
    }
    if (arg === "--stdin") {
      options.stdin = true;
      continue;
    }
    if (arg === "-v" || arg === "--verbose") {
      options.verbose = true;
      continue;
    }
    if (arg === "-q" || arg === "--quiet") {
      options.quiet = true;
      continue;
    }
    if (arg === "-h" || arg === "--help") {
      options.help = true;
      continue;
    }
    if (arg.startsWith("--max-markdown-chars=")) {
      const value = Number(arg.split("=")[1] ?? "");
      options.maxMarkdownChars = value > 0 ? value : null;
      continue;
    }
    if (arg === "--max-markdown-chars" && argv[index + 1]) {
      const value = Number(argv[index + 1] ?? "");
      options.maxMarkdownChars = value > 0 ? value : null;
      index += 1;
    }
  }

  return options;
}

function showHelp(): void {
  console.log(
    [
      "Usage: em2md [options]",
      "",
      "Options:",
      "  --input-dir DIR",
      "  --output-dir DIR",
      "  --done-dir DIR",
      "  --newest-first",
      "  --keep-input",
      "  --stdin",
      "  --max-markdown-chars N",
      "  -v, --verbose",
      "  -q, --quiet",
      "  -h, --help",
    ].join("\n"),
  );
}

export function truncateMarkdown(markdown: string, maxChars?: number | null): string {
  if (!maxChars || maxChars <= 0 || markdown.length <= maxChars) {
    return markdown;
  }

  const suffix = "\n\n... (truncated)";
  const allowed = Math.max(0, maxChars - suffix.length);
  return markdown.slice(0, allowed) + suffix;
}

function createCliLogger(options: CliOptions): RequiredLogger {
  const debugEnabled = options.verbose && !options.quiet;
  const infoEnabled = !options.quiet;

  return {
    info: (message) => {
      if (infoEnabled) {
        console.log(`${timestamp()} - INFO - ${message}`);
      }
    },
    warn: (message) => {
      console.warn(`${timestamp()} - WARNING - ${message}`);
    },
    debug: (message) => {
      if (debugEnabled) {
        console.log(`${timestamp()} - DEBUG - ${message}`);
      }
    },
    error: (message) => {
      console.error(`${timestamp()} - ERROR - ${message}`);
    },
  };
}

function timestamp(): string {
  return new Date().toISOString().replace("T", " ").slice(0, 19);
}

function decodeEmailHeaderValue(value: string | undefined | null): string {
  if (!value) {
    return "";
  }

  return he.decode(value).replace(/\s+/g, " ").trim();
}

function decodeAddressField(
  value: { text?: string | null } | Array<{ text?: string | null }> | null | undefined,
): string {
  if (Array.isArray(value)) {
    return value.map((entry) => decodeEmailHeaderValue(entry?.text ?? "")).filter(Boolean).join(", ");
  }
  return decodeEmailHeaderValue(value?.text ?? "");
}

function bufferFromAttachmentContent(content: unknown): Buffer {
  if (Buffer.isBuffer(content)) {
    return content;
  }
  if (content instanceof Uint8Array) {
    return Buffer.from(content);
  }
  if (typeof content === "string") {
    return Buffer.from(content, "utf8");
  }
  return Buffer.alloc(0);
}

function cleanHtmlToText(html: string): string {
  const $ = cheerio.load(html);

  $("script, style, noscript, head, meta, link, title, svg, iframe, object, embed").remove();

  $("img").each((_, element) => {
    const node = $(element);
    const alt = node.attr("alt")?.trim() || node.attr("aria-label")?.trim() || node.attr("title")?.trim();
    if (alt) {
      node.replaceWith(alt);
      return;
    }
    node.remove();
  });

  $("a").each((_, element) => {
    const node = $(element);
    const text = node.text().trim();
    const imageAlt = node.find("img").first().attr("alt")?.trim() || "";
    const label = text || node.attr("aria-label")?.trim() || node.attr("title")?.trim() || imageAlt;

    if (label) {
      node.replaceWith(label);
      return;
    }
    node.remove();
  });

  const cleanedHtml = $.root().html() ?? html;
  return htmlToText(cleanedHtml, {
    wordwrap: false,
    selectors: [
      { selector: "script", format: "skip" },
      { selector: "style", format: "skip" },
      { selector: "noscript", format: "skip" },
      { selector: "svg", format: "skip" },
      { selector: "iframe", format: "skip" },
    ],
  });
}

function cleanTextBase(value: string): string {
  let text = he.decode(value ?? "");
  text = text.normalize("NFKC");
  text = text.replace(/[\u00A0\u1680\u2000-\u200A\u202F\u205F\u3000]/g, " ");
  text = text.replace(/\r\n?/g, "\n");
  text = text.replace(/[\p{Cf}\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/gu, "");
  text = text.split("\n").map((line) => line.replace(/[ \t]+$/g, "")).join("\n");
  return text;
}

function compactWhitespace(text: string): string {
  let next = text;
  next = next.replace(/[ \t]{2,}/g, " ");
  next = next.replace(/\n[ \t]+\n/g, "\n\n");
  next = next.replace(/\n{3,}/g, "\n\n");
  return next;
}

export function cleanText(value: string): string {
  let text = cleanTextBase(value);
  text = text.replace(/\([^\n)]*https?:\/\/[^\n)]*\)/gi, "");
  text = text.replace(/[<\[]?https?:\/\/[^\s>\]]+[>\]]?/gi, "");
  text = text.replace(/\(\s*\)/g, "");
  text = text.split("\n").filter((line) => {
    const trimmed = line.trim();
    if (!trimmed) {
      return true;
    }
    if (/^[<\[]?https?:\/\/[^\s>\]]+[>\]]?$/i.test(trimmed)) {
      return false;
    }
    return true;
  }).join("\n");
  text = text.replace(/^\s*[<>\[\]()]+\s*$/gm, "");
  text = compactWhitespace(text);
  return text.trim();
}

function cleanTextWithoutUrlStripping(value: string): string {
  const text = compactWhitespace(cleanTextBase(value));
  return text.trim();
}

function looksLikePseudoPlainHtml(text: string): boolean {
  if (!text.trim()) {
    return false;
  }

  const strongPatterns = [
    /<a\b[^>]*href\s*=/i,
    /<img\b/i,
    /<table\b/i,
    /<\/?(html|body)\b/i,
  ];

  if (strongPatterns.some((pattern) => pattern.test(text))) {
    return true;
  }

  const weakTagMatches = text.match(/<\/?(div|span|p|br|tr|td)\b[^>]*>/gi) ?? [];
  return weakTagMatches.length >= 3;
}

function measureBodyQuality(body: string): BodyQualityMetrics {
  const lines = body.split("\n");
  const nonEmptyLines = lines.filter((line) => line.trim()).length;
  const informativeLines = lines.filter((line) => {
    const trimmed = line.trim();
    if (!trimmed) {
      return false;
    }
    if (!/[A-Za-z\u4E00-\u9FFF]/.test(trimmed)) {
      return false;
    }
    if (/^\[?.*https?:\/\//i.test(trimmed) || /^https?:\/\//i.test(trimmed)) {
      return false;
    }
    if (/(rsc=<|campaignId|utm_|=3D|=0A)/i.test(trimmed) && trimmed.length < 120) {
      return false;
    }
    return true;
  }).length;

  return {
    chars: body.length,
    nonEmptyLines,
    informativeLines,
    noiseMarkers: (body.match(/rsc=<|campaignId|utm_|=3D|=0A/gi) ?? []).length,
  };
}

export function choosePreferredBody(
  htmlBody: string,
  textBody: string,
): { body: string; source: "html" | "text"; reason: string } {
  const htmlMetrics = measureBodyQuality(htmlBody);
  const textMetrics = measureBodyQuality(textBody);

  const htmlEmpty = htmlMetrics.nonEmptyLines === 0;
  const textEmpty = textMetrics.nonEmptyLines === 0;

  if (htmlEmpty && !textEmpty) {
    return { body: textBody, source: "text", reason: "html body empty" };
  }
  if (textEmpty && !htmlEmpty) {
    return { body: htmlBody, source: "html", reason: "text body empty" };
  }
  if (htmlEmpty && textEmpty) {
    return { body: "", source: "html", reason: "both bodies empty" };
  }
  if (!htmlEmpty && looksLikePseudoPlainHtml(textBody)) {
    return { body: htmlBody, source: "html", reason: "text body contains embedded html markup" };
  }

  const maxChars = Math.max(htmlMetrics.chars, textMetrics.chars, 1);
  const ratioDelta = Math.abs(htmlMetrics.chars - textMetrics.chars) / maxChars;

  if (ratioDelta <= CLOSE_LENGTH_RATIO_THRESHOLD) {
    const htmlInfoDensity = htmlMetrics.informativeLines / Math.max(htmlMetrics.nonEmptyLines, 1);
    const textInfoDensity = textMetrics.informativeLines / Math.max(textMetrics.nonEmptyLines, 1);

    if (htmlInfoDensity !== textInfoDensity) {
      return htmlInfoDensity > textInfoDensity
        ? { body: htmlBody, source: "html", reason: "close lengths; higher informative density" }
        : { body: textBody, source: "text", reason: "close lengths; higher informative density" };
    }

    if (htmlMetrics.informativeLines !== textMetrics.informativeLines) {
      return htmlMetrics.informativeLines > textMetrics.informativeLines
        ? { body: htmlBody, source: "html", reason: "close lengths; more informative lines" }
        : { body: textBody, source: "text", reason: "close lengths; more informative lines" };
    }

    const htmlNoiseDensity = htmlMetrics.noiseMarkers / Math.max(htmlMetrics.chars, 1);
    const textNoiseDensity = textMetrics.noiseMarkers / Math.max(textMetrics.chars, 1);

    if (htmlNoiseDensity !== textNoiseDensity) {
      return htmlNoiseDensity < textNoiseDensity
        ? { body: htmlBody, source: "html", reason: "close lengths; lower noise density" }
        : { body: textBody, source: "text", reason: "close lengths; lower noise density" };
    }
  }

  return htmlMetrics.chars <= textMetrics.chars
    ? { body: htmlBody, source: "html", reason: "shorter body" }
    : { body: textBody, source: "text", reason: "shorter body" };
}

function parseDateValue(value: Date | string | null | undefined): DateValue {
  if (!value) {
    return null;
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  const parsed = new Date(value);
  if (!Number.isNaN(parsed.getTime())) {
    return parsed;
  }
  return value;
}

function formatDateValue(value: DateValue): string {
  if (!value) {
    return "";
  }
  if (typeof value === "string") {
    return value;
  }

  const year = value.getFullYear();
  const month = String(value.getMonth() + 1).padStart(2, "0");
  const day = String(value.getDate()).padStart(2, "0");
  const hours = String(value.getHours()).padStart(2, "0");
  const minutes = String(value.getMinutes()).padStart(2, "0");
  const seconds = String(value.getSeconds()).padStart(2, "0");
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

function extractThreadParts(bodyText: string): EmailPart[] {
  const threadParts: EmailPart[] = [];
  const patterns = [
    /From:[\s]*(.*?)[\r\n]+Sent:[\s]*(.*?)[\r\n]+To:[\s]*(.*?)(?:[\r\n]+Cc:[\s]*(.*?))?[\r\n]+Subject:[\s]*(.*?)[\r\n]+/gims,
    /On[\s]*(.*?),[\s]*(.*?)[\s]+wrote:[\r\n]+/gims,
    /On[\s]*(.*?)[\s]+at[\s]+(.*?),[\s]*(.*?)[\s]+wrote:[\r\n]+/gims,
  ];

  for (const pattern of patterns) {
    const matches = [...bodyText.matchAll(pattern)];
    if (!matches.length) {
      continue;
    }

    for (const match of matches) {
      const emailPart: EmailPart = pattern.source.includes("From:")
        ? {
            from: match[1]?.trim() ?? "",
            date: match[2]?.trim() ?? "",
            to: match[3]?.trim() ?? "",
            cc: match[4]?.trim() ?? "",
            subject: match[5]?.trim() ?? "",
            body: "",
            attachments: [],
          }
        : {
            from: match[3]?.trim() ?? match[1]?.trim() ?? "",
            date: match[1] && match[2] ? `${match[1].trim()} ${match[2].trim()}` : "",
            to: "",
            cc: "",
            subject: "",
            body: "",
            attachments: [],
          };

      const nextBodyStart = (match.index ?? 0) + match[0].length;
      emailPart.body = cleanText(bodyText.slice(nextBodyStart).trim());
      threadParts.push(emailPart);
    }

    if (threadParts.length) {
      break;
    }
  }

  return threadParts;
}

function hashFeature(feature: string): bigint {
  const digest = createHash("sha256").update(feature).digest();
  return digest.subarray(0, 8).reduce((acc, byte) => (acc << 8n) | BigInt(byte), 0n);
}

function simhash(text: string, bits = 64): bigint {
  const normalized = text.normalize("NFKC").replace(/\s+/g, " ").toLowerCase().trim();
  const words = normalized ? normalized.split(" ") : [];
  const features = [...words, ...words.slice(0, -1).map((word, index) => `${word} ${words[index + 1]}`)];
  const votes = Array.from({ length: bits }, () => 0);

  for (const feature of features) {
    const hash = hashFeature(feature);
    for (let bit = 0; bit < bits; bit += 1) {
      const isSet = (hash >> BigInt(bit)) & 1n;
      votes[bit] += isSet ? 1 : -1;
    }
  }

  let fingerprint = 0n;
  for (let bit = 0; bit < bits; bit += 1) {
    if (votes[bit] > 0) {
      fingerprint |= 1n << BigInt(bit);
    }
  }

  return fingerprint;
}

function hammingDistance(left: bigint, right: bigint): number {
  let xor = left ^ right;
  let count = 0;

  while (xor !== 0n) {
    count += Number(xor & 1n);
    xor >>= 1n;
  }

  return count;
}

function emailFeatureHash(emailPart: EmailPart): bigint {
  let content = "";
  if (emailPart.from) {
    content += `${emailPart.from} `;
  }
  if (emailPart.subject) {
    content += `${emailPart.subject} `;
  }

  const lines = emailPart.body.split("\n").map((line) => line.trim()).filter(Boolean);
  content += lines.slice(0, 5).join(" ");
  return simhash(content);
}

function deduplicateEmails(emails: EmailPart[], threshold = 8): EmailPart[] {
  if (!emails.length) {
    return [];
  }

  const normalized = emails.map((emailPart) => {
    const copy = structuredClone(emailPart) as EmailPart;
    if (copy.date instanceof Date && !Number.isNaN(copy.date.getTime())) {
      copy.date = new Date(copy.date.getTime());
    }
    return copy;
  });

  const emailHashes = normalized.map((emailPart) => ({ emailPart, hash: emailFeatureHash(emailPart) }));
  emailHashes.sort((left, right) => {
    const leftDate = left.emailPart.date instanceof Date ? left.emailPart.date.getTime() : Number.NEGATIVE_INFINITY;
    const rightDate = right.emailPart.date instanceof Date ? right.emailPart.date.getTime() : Number.NEGATIVE_INFINITY;
    return rightDate - leftDate;
  });

  const uniqueEmails: EmailPart[] = [];
  const usedIndices = new Set<number>();

  for (let index = 0; index < emailHashes.length; index += 1) {
    if (usedIndices.has(index)) {
      continue;
    }

    uniqueEmails.push(emailHashes[index].emailPart);
    usedIndices.add(index);

    for (let candidate = 0; candidate < emailHashes.length; candidate += 1) {
      if (candidate === index || usedIndices.has(candidate)) {
        continue;
      }

      if (hammingDistance(emailHashes[index].hash, emailHashes[candidate].hash) <= threshold) {
        usedIndices.add(candidate);
      }
    }
  }

  return uniqueEmails;
}

function sanitizeFilename(filename: string): string {
  return filename.replace(/[^\w.-]/g, "_");
}

function normalizeAttachmentFilename(filename: string | undefined): string {
  const decoded = decodeEmailHeaderValue(filename ?? "");
  return decoded || "attachment.bin";
}

async function collectEmailsFromParsed(parsed: ParsedMailLike, logger: RequiredLogger): Promise<EmailPart[]> {
  const emails: EmailPart[] = [];

  const htmlBody = typeof parsed.html === "string" && parsed.html.trim()
    ? cleanText(cleanHtmlToText(parsed.html))
    : "";
  const textBody = typeof parsed.text === "string" && parsed.text.trim()
    ? cleanTextWithoutUrlStripping(parsed.text)
    : "";
  const selectedBody = choosePreferredBody(htmlBody, textBody);

  const mainEmail: EmailPart = {
    date: parseDateValue(parsed.date ?? null),
    from: decodeAddressField(parsed.from),
    to: decodeAddressField(parsed.to),
    cc: decodeAddressField(parsed.cc),
    subject: decodeEmailHeaderValue(parsed.subject ?? ""),
    body: selectedBody.body,
    attachments: [],
  };

  logger.debug(`Extracting email: Subject='${mainEmail.subject}', From='${mainEmail.from}'`);
  logger.debug(`Body selected from ${selectedBody.source}: ${selectedBody.reason}`);
  emails.push(mainEmail);

  for (const attachment of parsed.attachments ?? []) {
    const content = bufferFromAttachmentContent(attachment.content);
    const contentType = attachment.contentType?.toLowerCase() ?? "";
    const filename = normalizeAttachmentFilename(attachment.filename);
    const looksLikeNestedMessage = content.length > 0 && (contentType === "message/rfc822" || filename.toLowerCase().endsWith(".eml"));

    if (looksLikeNestedMessage) {
      try {
        const nestedParsed = await simpleParser(content);
        const nestedEmails = await collectEmailsFromParsed(nestedParsed, logger);
        emails.push(...nestedEmails);
        continue;
      } catch (error) {
        logger.debug(`Failed to parse nested message attachment: ${(error as Error).message}`);
      }
    }

    if (content.length > 0) {
      mainEmail.attachments.push({
        filename,
        content,
        contentType: attachment.contentType ?? "application/octet-stream",
      });
    }
  }

  if (mainEmail.body) {
    const threadParts = extractThreadParts(mainEmail.body);
    if (threadParts.length) {
      emails.push(...threadParts);
    }
  }

  return emails;
}

function createMarkdownContent(emails: EmailPart[], newestFirst: boolean): string {
  const sortedEmails = [...emails].sort((left, right) => {
    const leftDate = left.date instanceof Date ? left.date.getTime() : Number.NEGATIVE_INFINITY;
    const rightDate = right.date instanceof Date ? right.date.getTime() : Number.NEGATIVE_INFINITY;
    return newestFirst ? rightDate - leftDate : leftDate - rightDate;
  });

  let markdown = "# Email Thread\n\n";

  sortedEmails.forEach((emailPart, index) => {
    markdown += `## Email ${index + 1}\n`;
    if (emailPart.date) {
      markdown += `- **Date**: ${formatDateValue(emailPart.date)}\n`;
    }
    markdown += `- **From**: ${emailPart.from}\n`;
    markdown += `- **To**: ${emailPart.to}\n`;
    if (emailPart.cc) {
      markdown += `- **CC**: ${emailPart.cc}\n`;
    }
    markdown += `- **Subject**: ${emailPart.subject}\n\n`;
    markdown += "### Content\n";
    markdown += `${emailPart.body}\n`;

    if (emailPart.attachments.length) {
      markdown += "\n### Attachments\n";
      for (const attachment of emailPart.attachments) {
        markdown += `- [${attachment.filename}](${attachment.filename})\n`;
      }
    }

    markdown += "\n---\n\n";
  });

  return markdown;
}

export async function convertEml(
  input: Buffer | Uint8Array | string,
  options: ConvertOptions = {},
): Promise<ConvertResult> {
  const logger = normalizeLogger(options.logger);
  const rawBytes = typeof input === "string" ? Buffer.from(input, "utf8") : Buffer.from(input);
  const parsed = await simpleParser(rawBytes);
  const emails = await collectEmailsFromParsed(parsed, logger);
  logger.info(`Total emails found: ${emails.length}`);

  const uniqueEmails = deduplicateEmails(emails);
  const markdown = truncateMarkdown(
    createMarkdownContent(uniqueEmails, options.newestFirst ?? false),
    options.maxMarkdownChars ?? null,
  );

  return { markdown, emails: uniqueEmails };
}

export async function convertEmlFile(filePath: string, options: ConvertOptions = {}): Promise<ConvertResult> {
  const rawBytes = await readFile(filePath);
  return convertEml(rawBytes, options);
}

async function writeConvertedEmlFile(filePath: string, options: CliOptions, logger: RequiredLogger): Promise<string> {
  const basename = path.basename(filePath);
  logger.info(`Processing: ${basename}`);

  const outputDirName = path.parse(basename).name;
  const outputDirPath = path.join(options.outputDir, outputDirName);
  await mkdir(outputDirPath, { recursive: true });

  const { markdown, emails } = await convertEmlFile(filePath, {
    newestFirst: options.newestFirst,
    maxMarkdownChars: options.maxMarkdownChars ?? null,
    logger,
  });

  const usedFilenames = new Map<string, number>();
  for (const emailPart of emails) {
    const newAttachments: EmailAttachment[] = [];

    for (const attachment of emailPart.attachments) {
      const safeFilename = sanitizeFilename(attachment.filename);
      const occurrence = usedFilenames.get(safeFilename) ?? 0;
      usedFilenames.set(safeFilename, occurrence + 1);

      const uniqueFilename = occurrence === 0
        ? safeFilename
        : `${path.parse(safeFilename).name}_${occurrence}${path.parse(safeFilename).ext}`;

      newAttachments.push({
        ...attachment,
        filename: uniqueFilename,
      });
    }

    emailPart.attachments = newAttachments;
  }

  const markdownWithAttachmentNames = truncateMarkdown(
    createMarkdownContent(emails, options.newestFirst),
    options.maxMarkdownChars ?? null,
  );
  const markdownPath = path.join(outputDirPath, `${outputDirName}.md`);
  await writeFile(markdownPath, markdownWithAttachmentNames || markdown, "utf8");

  for (const emailPart of emails) {
    for (const attachment of emailPart.attachments) {
      await writeFile(path.join(outputDirPath, attachment.filename), attachment.content);
    }
  }

  if (!options.keepInput) {
    await mkdir(options.doneDir, { recursive: true });
    const donePath = path.join(options.doneDir, basename);
    await rename(filePath, donePath);
    logger.info(`Moving processed file to: ${donePath}`);
  }

  logger.info(`Successfully processed: ${basename}`);
  return markdownPath;
}

async function readStdin(): Promise<Buffer> {
  const chunks: Buffer[] = [];
  for await (const chunk of process.stdin) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }
  return Buffer.concat(chunks);
}

export async function runCli(argv = process.argv.slice(2)): Promise<void> {
  const options = parseArgs(argv);
  if (options.help) {
    showHelp();
    return;
  }

  const logger = createCliLogger(options);

  if (options.stdin) {
    const rawBytes = await readStdin();
    if (!rawBytes.length) {
      logger.warn("No stdin content received");
      return;
    }

    const { markdown } = await convertEml(rawBytes, {
      newestFirst: options.newestFirst,
      maxMarkdownChars: options.maxMarkdownChars ?? null,
      logger,
    });
    process.stdout.write(markdown);
    return;
  }

  logger.info("Starting EML to Markdown converter");
  logger.info(
    `Settings: newest_first=${options.newestFirst}, input_dir=${options.inputDir}, output_dir=${options.outputDir}, max_markdown_chars=${options.maxMarkdownChars ?? "unlimited"}`,
  );

  for (const dir of [options.inputDir, options.outputDir, options.doneDir]) {
    await mkdir(dir, { recursive: true });
  }

  const entries = await readdir(options.inputDir, { withFileTypes: true });
  const emlFiles = entries.filter((entry) => entry.isFile() && entry.name.toLowerCase().endsWith(".eml"));

  if (!emlFiles.length) {
    logger.warn(`No EML files found in '${options.inputDir}' directory`);
    return;
  }

  logger.info(`Found ${emlFiles.length} EML file(s) to process`);
  logger.info("=".repeat(60));

  const processedFiles: Array<{ original: string; converted: string }> = [];
  const failedFiles: Array<{ original: string; error: string }> = [];

  for (let index = 0; index < emlFiles.length; index += 1) {
    const entry = emlFiles[index];
    const filePath = path.join(options.inputDir, entry.name);
    logger.info(`[${index + 1}/${emlFiles.length}] Starting: ${entry.name}`);

    try {
      const converted = await writeConvertedEmlFile(filePath, options, logger);
      processedFiles.push({ original: entry.name, converted });
      logger.info(`[${index + 1}/${emlFiles.length}] Success: ${entry.name}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      logger.error(`[${index + 1}/${emlFiles.length}] Failed: ${entry.name} - ${message}`);
      failedFiles.push({ original: entry.name, error: message });
    }

    logger.info("-".repeat(60));
  }

  logger.info("=".repeat(60));
  logger.info("CONVERSION SUMMARY");
  logger.info("=".repeat(60));
  logger.info(`Total files found: ${emlFiles.length}`);
  logger.info(`Successfully processed: ${processedFiles.length}`);
  logger.info(`Failed: ${failedFiles.length}`);

  if (processedFiles.length) {
    logger.info("");
    logger.info("Successfully converted:");
    for (const item of processedFiles) {
      logger.info(`  ${item.original} -> ${item.converted}`);
    }
  }

  if (failedFiles.length) {
    logger.warn("");
    logger.warn("Failed to convert:");
    for (const item of failedFiles) {
      logger.warn(`  ${item.original}: ${item.error}`);
    }
  }
}
