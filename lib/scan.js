// Finds transcript download links in a JSON API response.
// Shared by inject.js (page) and the node tests. No dependencies.
(function (root) {
  'use strict';

  const MAX_DEPTH = 12;

  // Is this object (found at `path`) a transcript entry?
  function looksLikeTranscript(obj, path) {
    if (typeof obj.temporaryDownloadUrl !== 'string') return false;
    if (!/^https:\/\//i.test(obj.temporaryDownloadUrl)) return false;
    if (path.some((k) => /transcript/i.test(k))) return true;
    if ('transcriptType' in obj) return true;
    return /transcript/i.test(obj.temporaryDownloadUrl.split('?')[0]);
  }

  function str(v) {
    return typeof v === 'string' && v ? v : undefined;
  }

  // Returns [{ name, language, transcriptType, url, itemName, createdDateTime, meetingDate }]
  function scanForTranscripts(json) {
    const found = [];
    if (!json || typeof json !== 'object') return found;
    const rootItemName = str(json.name);
    const rootCreated = str(json.createdDateTime);

    (function walk(node, path, parentItem) {
      if (!node || typeof node !== 'object' || path.length > MAX_DEPTH) return;
      if (Array.isArray(node)) {
        node.forEach((child, i) => walk(child, path.concat(String(i)), parentItem));
        return;
      }
      // Track the nearest object that looks like a drive item (has name + createdDateTime).
      const item = str(node.name) && str(node.createdDateTime) ? node : parentItem;
      if (looksLikeTranscript(node, path)) {
        found.push({
          name: str(node.displayName) || str(node.name) || str(node.fileName) || (item && str(item.name)) || rootItemName || 'transcript',
          language: str(node.languageTag) || str(node.language) || str(node.locale),
          transcriptType: str(node.transcriptType),
          url: node.temporaryDownloadUrl,
          itemName: (item && item !== node && str(item.name)) || rootItemName,
          createdDateTime: str(node.createdDateTime) || (item && str(item.createdDateTime)) || rootCreated,
          meetingDate: str(node.meetingStartDateTime) || str(json.meetingStartDateTime) || (json.media && str(json.media.meetingStartDateTime)),
        });
      }
      for (const key of Object.keys(node)) {
        const v = node[key];
        if (v && typeof v === 'object') walk(v, path.concat(key), item);
      }
    })(json, [], undefined);

    // De-duplicate by URL.
    const seen = new Set();
    return found.filter((t) => (seen.has(t.url) ? false : seen.add(t.url)));
  }

  // Cheap pre-check so we only JSON.parse responses that can contain a link.
  function mightContainTranscript(text) {
    return typeof text === 'string' && text.indexOf('temporaryDownloadUrl') !== -1;
  }

  const api = { scanForTranscripts, mightContainTranscript };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else Object.defineProperty(root, '__teamsTranscriptScan', { value: api, configurable: false });
})(typeof globalThis !== 'undefined' ? globalThis : this);
