import {
  Document,
  ExternalHyperlink,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun
} from "./lib/docx-lite.js";

function toRun(run) {
  const text = run.text || "";
  return new TextRun({
    text,
    bold: !!run.bold,
    italic: !!run.italic,
    underline: !!run.underline
  });
}

function paragraphChildrenFromRuns(runs = []) {
  const children = [];
  let pendingLink = null;
  let pendingRuns = [];

  function flushLink() {
    if (!pendingLink) return;
    children.push(new ExternalHyperlink({ link: pendingLink, children: pendingRuns }));
    pendingLink = null;
    pendingRuns = [];
  }

  for (const run of runs) {
    if (run.href) {
      if (pendingLink !== run.href) {
        flushLink();
        pendingLink = run.href;
      }
      pendingRuns.push(toRun(run));
      continue;
    }
    flushLink();
    children.push(toRun(run));
  }
  flushLink();
  return children;
}

function paragraphFromRuns(runs, options = {}) {
  return new Paragraph({
    headingLevel: options.headingLevel || null,
    numbering: options.numbering || null,
    children: paragraphChildrenFromRuns(runs)
  });
}

function blocksToDocElements(blocks, state = { listLevel: 0 }) {
  const out = [];
  for (const block of blocks || []) {
    if (!block) continue;

    if (block.type === "heading") {
      out.push(paragraphFromRuns(block.runs || [], { headingLevel: Math.min(6, Math.max(1, block.level || 1)) }));
      continue;
    }

    if (block.type === "paragraph") {
      out.push(paragraphFromRuns(block.runs || []));
      continue;
    }

    if (block.type === "list") {
      const numId = block.ordered ? 2 : 1;
      for (const item of block.items || []) {
        const childBlocks = item.blocks || [];
        let emittedFirst = false;
        for (const itemBlock of childBlocks) {
          if (itemBlock.type === "paragraph" || itemBlock.type === "heading") {
            const runs = itemBlock.runs || [];
            out.push(
              paragraphFromRuns(runs, {
                headingLevel: itemBlock.type === "heading" ? Math.min(6, Math.max(1, itemBlock.level || 1)) : null,
                numbering: { numId, level: state.listLevel || 0 }
              })
            );
            emittedFirst = true;
          } else if (itemBlock.type === "list") {
            out.push(...blocksToDocElements([itemBlock], { listLevel: (state.listLevel || 0) + 1 }));
          } else if (itemBlock.type === "table") {
            out.push(...blocksToDocElements([itemBlock], state));
          }
        }
        if (!emittedFirst && childBlocks.length === 0) {
          out.push(paragraphFromRuns([{ text: "" }], { numbering: { numId, level: state.listLevel || 0 } }));
        }
      }
      continue;
    }

    if (block.type === "table") {
      const rows = [];
      for (const row of block.rows || []) {
        const cells = [];
        for (const cell of row || []) {
          const cellChildren = blocksToDocElements(cell.blocks || []);
          const paragraphs = cellChildren.filter((el) => el instanceof Paragraph);
          cells.push(new TableCell({ children: paragraphs.length ? paragraphs : [new Paragraph()] }));
        }
        if (cells.length) rows.push(new TableRow({ children: cells }));
      }
      if (rows.length) out.push(new Table({ rows }));
    }
  }
  return out;
}

export function safeFilename(baseTitle) {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const timestamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}_${pad(now.getHours())}-${pad(now.getMinutes())}`;
  const slug = (baseTitle || "document")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80) || "document";
  return `${slug}_${timestamp}.docx`;
}

export function buildDocxBlob(payload) {
  const title = payload && payload.title ? payload.title : "Untitled";
  const children = [];
  children.push(
    new Paragraph({
      headingLevel: 1,
      children: [new TextRun({ text: title, bold: true })]
    })
  );
  children.push(...blocksToDocElements((payload && payload.blocks) || []));

  const documentModel = new Document({
    title,
    children
  });
  return Packer.toBlob(documentModel);
}
