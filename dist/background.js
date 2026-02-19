import { buildDocxBlob, safeFilename } from "./doc-builder.js";

async function getActiveTab() {
  const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  return tabs[0];
}

async function ensureContentScript(tabId) {
  await chrome.scripting.executeScript({
    target: { tabId },
    files: ["dist/content.js"]
  });
}

function sendExtractRequest(tabId) {
  return chrome.tabs.sendMessage(tabId, { type: "TODOCX_EXTRACT" });
}

async function exportActiveTabToDocx(tab) {
  if (!tab || !tab.id) return;

  await ensureContentScript(tab.id);
  const response = await sendExtractRequest(tab.id);
  if (!response || !response.ok || !response.data) {
    throw new Error((response && response.error) || "Failed to extract content from page.");
  }

  const blob = buildDocxBlob(response.data);
  const filename = safeFilename(response.data.title || tab.title || "document");
  const url = await blobToDataUrl(blob);
  await chrome.downloads.download({
    url,
    filename,
    saveAs: false
  });
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("Failed to convert Blob to data URL."));
    reader.readAsDataURL(blob);
  });
}

chrome.action.onClicked.addListener(async (tab) => {
  try {
    const activeTab = tab && tab.id ? tab : await getActiveTab();
    await exportActiveTabToDocx(activeTab);
  } catch (error) {
    console.error("ToDOCX export error:", error);
  }
});
