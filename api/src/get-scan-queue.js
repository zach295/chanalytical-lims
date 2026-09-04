const { app } = require('@azure/functions');
const { listItems, LISTS } = require('../shared/graph');

app.http('get-scan-queue', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      // ── Pending scans from Review Queue ──────────────────────────────────────
      const queueItems = await listItems(LISTS.REVIEW_QUEUE, { top: 100 });
      const filtered = queueItems.filter(r => {
        const status = r.Title || r.ReviewStatus || '';
        return status !== 'Approved' && status !== 'Discarded' && status !== 'Processed';
      });

      // ── Cross-reference: remove any pending item whose FileID is in Archived Intake ──
      // When a scan is approved, approve-scan.js stores FileID in Archived Intake field_15.
      // This is the server-side safety net in case the Review Queue row wasn't deleted.
      let approvedFileIds = new Set();
      try {
        const aiItems = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 });
        aiItems.forEach(r => { if (r.field_15) approvedFileIds.add(String(r.field_15).trim()); });
      } catch(e) { context.log('[get-scan-queue] Archived Intake cross-ref failed:', e.message); }

      const parseTests = (value) => {
        const raw = String(value || '').trim();
        if (!raw) return [];
        // New rows are pipe/semicolon delimited so commas inside test names are preserved.
        if (/[|;]/.test(raw)) {
          return raw.split(/\s*(?:\||;)\s*/).map(t => t.trim()).filter(Boolean);
        }
        // Backward compatibility for existing rows written comma-delimited.
        // Do not split commas that are part of names like "Iron, Total" or "Arsenic, Speciation".
        return raw.split(/,\s+(?!(?:Total|Speciation)\b)/i).map(t => t.trim()).filter(Boolean);
      };

      const pending = filtered
        .filter(r => !approvedFileIds.has(String(r.FileID || r.FileId || '').trim()))
        .map(r => ({
        fileId:           r.FileID       || r.FileId       || '',
        barcodeId:        r.BarcodeID    || r.BarcodeId    || '',
        baseId:           (r.BarcodeID    || r.BarcodeId    || '').split(' ')[0].trim(),
        customer:         r.ClientName   || '',
        email:            r.Email        || '',
        dateDrawn:        r.SampleDate   || '',
        timeDrawn:        r.SampleTime   || '',
        receivedDate:     r.ReceivedDate || '',
        receivedTime:     r.ReceivedTime || '',
        location:         r.Address      || '',
        city:             r.City         || '',
        state:            r.State        || 'ME',
        zip:              r.Zip          || '',
        services:         r.TestSelections || '',
        tests:            parseTests(r.TestSelections),
        confidence:       r.OCRConfidence || 0,
        processedDate:    r.ProcessedDate || '',
        reviewStatus:     r.Title        || r.ReviewStatus || 'Pending',
        validationErrors: r.ValidationErrors || '',
        waterType:        r.WaterType    || '',
        phone:            r.Phone        || '',
        billingAddress:   r.BillingAddress || '',
        isNewClient:      r.IsNewClient  === 'Yes',
        formType:         r.FormType     || 'public',
        fileName:         r.FileName     || '',
        _rowIndex:        r._id,
        _ocrDebug:        r.OCRDebug || '',
      }));

      // ── Recently approved from Archived Intake ────────────────────────────────
      const archivedItems = await listItems(LISTS.ARCHIVED_INTAKE, { top: 200 });

      archivedItems.sort((a, b) => {
        const da = new Date(a.Title || 0);
        const db = new Date(b.Title || 0);
        return db - da;
      });

      const groupedByTs = {};
      archivedItems.forEach(r => {
        const ts = r.Title || '';
        if (!ts) return;
        if (!groupedByTs[ts]) {
          groupedByTs[ts] = {
            ts,
            labIds:     [],
            coaTests:   [],
            customer:   r.field_3  || '',
            approvedBy: r.field_12 || '',
          };
        }
        if (r.field_1) groupedByTs[ts].labIds.push(r.field_1);
        if (r.field_2) groupedByTs[ts].coaTests.push(r.field_2);
      });

      const allSorted = Object.values(groupedByTs).sort((a, b) => {
        const tsA = new Date(a.ts || 0);
        const tsB = new Date(b.ts || 0);
        return tsB - tsA;
      });

      const today = new Date().toLocaleDateString('en-CA', { timeZone: 'America/New_York' });
      const todayApproved = allSorted.slice(0, 10).map(g => ({
        ts:         g.ts,
        labIds:     [...new Set(g.labIds)],
        tests:      [...new Set(g.coaTests)],
        customer:   g.customer,
        approvedBy: g.approvedBy,
      }));
      const todayCount = allSorted.filter(g => g.ts.startsWith(today)).length;

      return {
        status: 200,
        headers: {
          'content-type': 'application/json',
          'cache-control': 'no-store, no-cache, must-revalidate',
          'pragma': 'no-cache',
        },
        body: JSON.stringify({ pending, todayApproved, todayCount }),
      };

    } catch(e) {
      context.log('[get-scan-queue] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
