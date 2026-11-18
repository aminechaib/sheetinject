// SheetInject Background Service Worker (Manifest V3)
chrome.action.onClicked.addListener(async (tab) => {
  if (!tab.id) return;

  try {
    const response = await chrome.tabs.sendMessage(tab.id, { type: 'SHEETINJECT_PING' });
    if (response && response.mounted) {
      await chrome.tabs.sendMessage(tab.id, { type: 'SHEETINJECT_TOGGLE' });
      return;
    }
  } catch (_err) {
    // Content script not injected in this tab yet
  }

  try {
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files: ['lib/xlsx.mini.min.js', 'content.js']
    });
  } catch (err) {
    console.error('SheetInject failed to initialize on this tab:', err);
  }
});
