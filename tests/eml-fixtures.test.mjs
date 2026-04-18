import test from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import { readdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";

import { convertEmlFile } from "../dist/index.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const fixturesDir = path.resolve(__dirname, "../fixtures");

test("all fixture .eml files can be converted", async () => {
  const entries = await readdir(fixturesDir, { withFileTypes: true });
  const emlFiles = entries
    .filter((entry) => entry.isFile() && entry.name.toLowerCase().endsWith(".eml"))
    .map((entry) => entry.name)
    .sort();

  assert.ok(emlFiles.length > 0, "expected at least one .eml fixture");

  for (const filename of emlFiles) {
    const fullPath = path.join(fixturesDir, filename);
    const result = await convertEmlFile(fullPath);

    assert.ok(result.markdown.startsWith("# Email Thread"), `${filename}: markdown should have thread heading`);
    assert.ok(result.emails.length > 0, `${filename}: expected at least one parsed email`);
    assert.ok(
      result.emails.some((email) => email.body.trim().length > 0),
      `${filename}: expected at least one email with non-empty body`,
    );
  }
});

test("002.eml respects maxMarkdownChars=200", async () => {
  const fullPath = path.join(fixturesDir, "002.eml");
  const result = await convertEmlFile(fullPath, { maxMarkdownChars: 200 });

  assert.equal(result.markdown.length, 200, "002.eml: markdown should be limited to exactly 200 chars");
  assert.ok(
    result.markdown.endsWith("\n\n... (truncated)"),
    "002.eml: markdown should include truncation suffix",
  );
});

test("fixture metadata fields are parsed correctly", async () => {
  const expected = {
    "001.eml": {
      from: '"ZDNet Announcements" <Online#3.19690.d9-iS5vl-4yBSBWq9RR.1@newsletter.online.com>',
      to: "qqqqqqqqqq-zdnet@spamassassin.taint.org",
      cc: "",
      subject: "Get the Perfect Mix of ZDNet's Best Stuff!",
      dateIso: "2002-07-10T23:05:42.000Z",
    },
    "002.eml": {
      from: "update@list.theregister.co.uk",
      to: '"Reg Subscribers" <update@list.theregister.co.uk>',
      cc: "",
      subject: "Reg Headlines Friday July 12",
      dateIso: "2002-07-12T02:00:01.000Z",
    },
    "003.eml": {
      from: '"joyful" <abcusa88@hotmail.com>',
      to: '"One Income Living" <OneIncomeLiving@groups.msn.com>',
      cc: "",
      subject: "Re: Hi! I'm new here.",
      dateIso: "2002-08-10T03:09:02.000Z",
    },
    "004.eml": {
      from: '"Media Unspun" <guterman@mediaunspun.imakenews.net>',
      to: "zzz-unspun@spamassassin.taint.org",
      cc: "",
      subject: "Brother, Can You Spare a Jet?",
      dateIso: "2002-08-12T13:11:31.000Z",
    },
    "005.eml": {
      from: "piro-test@clear-code.com",
      to: "piro.outsider.reflex+1@gmail.com, piro.outsider.reflex+2@gmail.com, mailmaster@example.com, mailmaster@example.org, webmaster@example.com, webmaster@example.org, webmaster@example.jp, mailmaster@example.jp",
      cc: "",
      subject: "test confirmation",
      dateIso: "2019-08-15T05:54:37.000Z",
    },
    "006.eml": {
      from: '"Ninon" <ussrivjuta@hotmail.com>',
      to: '"Ambrose" <paliourg@iit.demokritos.gr>',
      cc: "",
      subject: "::: HOODIA CACTUS MAKES YOU LOSE WEIGHT FOREVER :::",
      dateIso: "2004-06-12T01:00:21.000Z",
    },
  };

  for (const [filename, meta] of Object.entries(expected)) {
    const fullPath = path.join(fixturesDir, filename);
    const result = await convertEmlFile(fullPath);
    const firstEmail = result.emails[0];

    assert.ok(firstEmail, `${filename}: expected at least one parsed email`);
    assert.equal(firstEmail.from, meta.from, `${filename}: from mismatch`);
    assert.equal(firstEmail.to, meta.to, `${filename}: to mismatch`);
    assert.equal(firstEmail.cc, meta.cc, `${filename}: cc mismatch`);
    assert.equal(firstEmail.subject, meta.subject, `${filename}: subject mismatch`);

    const parsedDateIso = firstEmail.date instanceof Date ? firstEmail.date.toISOString() : String(firstEmail.date);
    assert.equal(parsedDateIso, meta.dateIso, `${filename}: date mismatch`);
  }
});

test("005.eml renders attachments in markdown", async () => {
  const fullPath = path.join(fixturesDir, "005.eml");
  const result = await convertEmlFile(fullPath);

  const attachments = result.emails.flatMap((email) => email.attachments);
  const attachmentNames = attachments.map((attachment) => attachment.filename).sort();

  assert.equal(attachments.length, 2, "005.eml: expected exactly 2 attachments");
  assert.deepEqual(
    attachmentNames,
    ["manifest.json", "sha1hash.txt"],
    "005.eml: attachment filenames mismatch",
  );

  assert.ok(result.markdown.includes("### Attachments"), "005.eml: markdown should include attachments section");
  assert.ok(
    result.markdown.includes("- [sha1hash.txt](sha1hash.txt)"),
    "005.eml: markdown should render sha1hash.txt attachment link",
  );
  assert.ok(
    result.markdown.includes("- [manifest.json](manifest.json)"),
    "005.eml: markdown should render manifest.json attachment link",
  );
});
