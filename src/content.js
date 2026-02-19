(() => {
  const DEBUG_NODES = true;
  // Add selectors here to treat matched elements as bold/italic content.
  const BOLD_STYLE_SELECTORS = ["b", "strong", "span.selection-bold"];
  const ITALIC_STYLE_SELECTORS = ["i", "em"];

  const BLOCK_TAGS = new Set([
    "ADDRESS",
    "ARTICLE",
    "ASIDE",
    "BLOCKQUOTE",
    "DIV",
    "DL",
    "DETAILS",
    "FIELDSET",
    "FIGCAPTION",
    "FIGURE",
    "H1",
    "H2",
    "H3",
    "H4",
    "H5",
    "H6",
    "HR",
    "LI",
    "MAIN",
    "NAV",
    "OL",
    "P",
    "PRE",
    "SECTION",
    "SUMMARY",
    "TABLE",
    "TBODY",
    "TD",
    "TFOOT",
    "TH",
    "THEAD",
    "TR",
    "UL"
  ]);

  // Entries can be plain tags ("script") or full CSS selectors
  // ("header.site-header", "#cookie-banner", "div.ad").
  const SKIP_MATCHERS = [
    "canvas",
    "footer",
    "form",
    "header",
    "iframe",
    "input",
    "noscript",
    "script",
    "select",
    "style",
    "svg",
    "textarea",
    "div#isi-component",
    "div.gsk-patient-grid",
    "div[data-component='medwatch']"
  ];
  const SKIP_TAGS = new Set(
    SKIP_MATCHERS.filter((entry) => /^[a-z][a-z0-9-]*$/i.test(entry)).map((entry) => entry.toUpperCase())
  );
  const SKIP_SELECTORS = SKIP_MATCHERS.filter((entry) => !/^[a-z][a-z0-9-]*$/i.test(entry));

  function isElement(node) {
    return node && node.nodeType === Node.ELEMENT_NODE;
  }

  function describeNode(node) {
    if (!node) return "<null>";
    if (node.nodeType === Node.TEXT_NODE) {
      const value = (node.nodeValue || "").replace(/\s+/g, " ").trim();
      return `#text("${value.slice(0, 80)}")`;
    }
    if (!isElement(node)) return `<node type=${node.nodeType}>`;
    const id = node.id ? `#${node.id}` : "";
    const className =
      typeof node.className === "string" && node.className.trim()
        ? `.${node.className.trim().replace(/\s+/g, ".")}`
        : "";
    return `<${node.tagName.toLowerCase()}${id}${className}>`;
  }

  function debugNode(prefix, node) {
    if (!DEBUG_NODES) return;
    console.debug(`[ToDOCX] ${prefix}: ${describeNode(node)}`, node);
  }

  function isInsideHeaderFooter(node) {
    if (!node) return false;
    const el = isElement(node) ? node : node.parentElement;
    return !!(el && el.closest("header, footer"));
  }

  function isHiddenByCss(el) {
    if (!el || !isElement(el)) return false;
    const style = window.getComputedStyle(el);
    if (!style) return false;
    if (style.display === "none" || style.visibility === "hidden" || style.visibility === "collapse") {
      return true;
    }

    // Only display:none on ancestors is guaranteed to hide descendants.
    let cursor = el.parentElement;
    while (cursor) {
      const parentStyle = window.getComputedStyle(cursor);
      if (parentStyle && parentStyle.display === "none") return true;
      cursor = cursor.parentElement;
    }
    return false;
  }

  function isSkippableElement(el) {
    return (
      SKIP_TAGS.has(el.tagName) ||
      matchesAnySelector(el, SKIP_SELECTORS) ||
      isInsideHeaderFooter(el) ||
      isHiddenByCss(el)
    );
  }

  function cleanText(text) {
    return text.replace(/\s+/g, " ");
  }

  function matchesAnySelector(el, selectors) {
    if (!el || !isElement(el)) return false;
    for (const selector of selectors) {
      if (!selector) continue;
      try {
        if (el.matches(selector)) return true;
      } catch (_err) {
        // Ignore invalid selectors to keep extraction resilient.
      }
    }
    return false;
  }

  function appendRun(runs, run) {
    if (!run || !run.text) return;
    if (!runs.length) {
      runs.push(run);
      return;
    }
    const prev = runs[runs.length - 1];
    const sameStyle =
      !!prev.bold === !!run.bold &&
      !!prev.italic === !!run.italic &&
      !!prev.underline === !!run.underline &&
      (prev.href || "") === (run.href || "");
    if (sameStyle) {
      prev.text += run.text;
    } else {
      runs.push(run);
    }
  }

  function collectInlineRuns(node, marks = {}, out = []) {
    debugNode("collectInlineRuns", node);
    if (!node) return out;

    if (node.nodeType === Node.TEXT_NODE) {
      const text = cleanText(node.nodeValue || "");
      if (text.trim()) {
        appendRun(out, {
          text,
          bold: !!marks.bold,
          italic: !!marks.italic,
          underline: !!marks.underline,
          href: marks.href || undefined
        });
      }
      return out;
    }

    if (!isElement(node) || isSkippableElement(node)) {
      return out;
    }

    const tag = node.tagName;
    if (tag === "BR") {
      appendRun(out, { text: "\n", ...marks });
      return out;
    }
    if (tag === "IMG") {
      const source = (node.getAttribute("src") || node.currentSrc || "").trim();
      if (source) {
        appendRun(out, { text: `[image]${source}[/image]` });
      }
      return out;
    }

    const nextMarks = {
      bold: marks.bold || matchesAnySelector(node, BOLD_STYLE_SELECTORS),
      italic: marks.italic || matchesAnySelector(node, ITALIC_STYLE_SELECTORS),
      underline: marks.underline || tag === "U",
      href: marks.href
    };
    if (tag === "A") {
      const href = node.getAttribute("href");
      if (href) nextMarks.href = href;
    }

    for (const child of node.childNodes) {
      collectInlineRuns(child, nextMarks, out);
    }

    return out;
  }

  function firstTextFromNode(node) {
    if (!node) return "";
    const text = cleanText(node.textContent || "").trim();
    return text;
  }

  function parseContainerChildren(container) {
    const blocks = [];
    let inlineBucket = [];

    function flushInlineBucket() {
      if (!inlineBucket.length) return;
      const text = inlineBucket.map((n) => firstTextFromNode(n)).join(" ").trim();
      if (!text) {
        inlineBucket = [];
        return;
      }
      const runs = [];
      for (const n of inlineBucket) collectInlineRuns(n, {}, runs);
      if (runs.length) blocks.push({ type: "paragraph", runs });
      inlineBucket = [];
    }

    for (const child of container.childNodes) {
      debugNode("collectInlineRuns", child);
      if (child.nodeType === Node.TEXT_NODE) {
        if ((child.nodeValue || "").trim()) inlineBucket.push(child);
        continue;
      }
      if (!isElement(child) || isSkippableElement(child)) continue;
      if (BLOCK_TAGS.has(child.tagName)) {
        flushInlineBucket();
        blocks.push(...parseBlock(child));
      } else {
        inlineBucket.push(child);
      }
    }

    flushInlineBucket();
    return blocks;
  }

  function parseList(el, ordered) {
    const items = [];
    for (const child of el.children) {
      if (!isElement(child) || child.tagName !== "LI" || isSkippableElement(child)) continue;
      const blocks = parseContainerChildren(child);
      if (!blocks.length) {
        const runs = collectInlineRuns(child);
        if (runs.length) blocks.push({ type: "paragraph", runs });
      }
      if (blocks.length) items.push({ blocks });
    }
    return items.length ? [{ type: "list", ordered, items }] : [];
  }

  function parseTable(el) {
    const rows = [];
    const trList = el.querySelectorAll(":scope > thead > tr, :scope > tbody > tr, :scope > tfoot > tr, :scope > tr");
    for (const tr of trList) {
      if (isSkippableElement(tr)) continue;
      const row = [];
      for (const cell of tr.children) {
        if (!isElement(cell)) continue;
        if (cell.tagName !== "TD" && cell.tagName !== "TH") continue;
        if (isSkippableElement(cell)) continue;
        const blocks = parseContainerChildren(cell);
        if (!blocks.length) {
          const runs = collectInlineRuns(cell);
          if (runs.length) blocks.push({ type: "paragraph", runs });
        }
        row.push({ blocks: blocks.length ? blocks : [{ type: "paragraph", runs: [{ text: "" }] }] });
      }
      if (row.length) rows.push(row);
    }
    return rows.length ? [{ type: "table", rows }] : [];
  }

  function parseBlock(el) {
    debugNode("parseBlock", el);
    if (!el || !isElement(el) || isSkippableElement(el)) return [];
    const tag = el.tagName;

    if (/^H[1-6]$/.test(tag)) {
      const runs = collectInlineRuns(el);
      if (!runs.length) return [];
      return [{ type: "heading", level: Number(tag[1]), runs }];
    }

    if (tag === "P") {
      const runs = collectInlineRuns(el);
      return runs.length ? [{ type: "paragraph", runs }] : [];
    }

    if (tag === "UL") return parseList(el, false);
    if (tag === "OL") return parseList(el, true);
    if (tag === "TABLE") return parseTable(el);
    if (tag === "LI") {
      const blocks = parseContainerChildren(el);
      if (blocks.length) return blocks;
      const runs = collectInlineRuns(el);
      return runs.length ? [{ type: "paragraph", runs }] : [];
    }

    const childBlocks = parseContainerChildren(el);
    if (childBlocks.length) return childBlocks;

    const runs = collectInlineRuns(el);
    return runs.length ? [{ type: "paragraph", runs }] : [];
  }

  function candidateScore(el) {
    if (!el || isSkippableElement(el)) return 0;
    const text = cleanText(el.innerText || "");
    const len = text.trim().length;
    if (!len) return 0;
    const rect = el.getBoundingClientRect();
    const area = Math.max(1, rect.width * rect.height);
    return len + Math.log(area + 1);
  }

  function pickRoot() {
    const mains = [...document.querySelectorAll("div#spa-root")].filter((el) => !isSkippableElement(el));
    if (mains.length) {
      mains.sort((a, b) => candidateScore(b) - candidateScore(a));
      return mains[0];
    }

    const articles = [...document.querySelectorAll("article")].filter((el) => !isSkippableElement(el));
    if (articles.length) {
      articles.sort((a, b) => candidateScore(b) - candidateScore(a));
      return articles[0];
    }

    const all = [...document.body.querySelectorAll("*")].filter((el) => {
      if (isSkippableElement(el)) return false;
      if (el.childElementCount === 0) return false;
      return cleanText(el.innerText || "").trim().length > 0;
    });
    let best = document.body;
    let bestScore = 0;
    for (const el of all) {
      const score = candidateScore(el);
      if (score > bestScore) {
        bestScore = score;
        best = el;
      }
    }
    return best;
  }

  function extractStructuredContent() {
    const root = pickRoot();
    const blocks = parseBlock(root);
    return {
      title: document.title || "Untitled",
      url: location.href,
      blocks
    };
  }

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    if (!message || message.type !== "TODOCX_EXTRACT") return;
    try {
      sendResponse({ ok: true, data: extractStructuredContent() });
    } catch (error) {
      sendResponse({
        ok: false,
        error: error && error.message ? error.message : "Extraction failed"
      });
    }
    return true;
  });
})();
