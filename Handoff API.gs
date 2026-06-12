/**
 * HANDOFF API — token-protected web app for the local `npm run handoff` orchestrator.
 *
 * The bound project already uses doGet() for the Leaflet map (Map_Link.gs), so this API lives on
 * doPost() and does not interfere with it. Node calls it over HTTP (like sheet-lookup.js).
 *
 * One-time setup:
 *   1. Run setupHandoffApiToken('some-long-random-string') once (Apps Script editor) — or set the
 *      'HANDOFF_API_TOKEN' Script Property under Project Settings → Script Properties.
 *   2. Deploy → New deployment → Web app | Execute as: Me | Who has access: Anyone.
 *   3. Put the /exec URL and the same token in the local .handoff.env file.
 *
 * Actions (POST JSON body { token, action, fuzeIds? }):
 *   - "ready"    → { ready: [{fuze, site}, ...] }            (no drafts created)
 *   - "cqStatus" → { status: [{fuze, present}, ...] }        (is each CQ cloud-visible in Drive?)
 *   - "generate" → { created: <n>, missing: [labels...] }    (creates the handoff drafts)
 */

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) body = JSON.parse(e.postData.contents);

    var expected = PropertiesService.getScriptProperties().getProperty('HANDOFF_API_TOKEN');
    if (!expected || String(body.token) !== String(expected)) {
      return jsonOut_({ error: 'unauthorized' });
    }

    switch (body.action) {
      case 'ready':
        return jsonOut_(generateHandoffDrafts(true)); // { ready: [...] }

      case 'cqStatus': {
        var fuzeIds = Array.isArray(body.fuzeIds) ? body.fuzeIds : [];
        var folder = getCqFolder_();
        var files = [];
        if (folder) { var it = folder.getFiles(); while (it.hasNext()) files.push(it.next()); }
        var status = fuzeIds.map(function(f) {
          return { fuze: String(f).trim(), present: !!findCqFile_(f, files) };
        });
        return jsonOut_({ status: status, folderResolved: !!folder });
      }

      case 'generate':
        return jsonOut_(generateHandoffDrafts(false)); // { created, missing }

      default:
        return jsonOut_({ error: 'unknown action: ' + body.action });
    }
  } catch (err) {
    return jsonOut_({ error: String(err) });
  }
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Run once from the Apps Script editor to set the shared secret used by the orchestrator.
function setupHandoffApiToken(token) {
  if (!token) throw new Error('Pass a token string, e.g. setupHandoffApiToken("xyz123...").');
  PropertiesService.getScriptProperties().setProperty('HANDOFF_API_TOKEN', String(token));
  console.log('HANDOFF_API_TOKEN set.');
}
