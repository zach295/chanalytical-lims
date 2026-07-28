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

      const pending = filtered.map(r => ({
        fileId:           r.FileID       || r.FileId       || r.fileId       || '',
        barcodeId:        r.BarcodeID    || r.BarcodeId    || r.barcodeId    || '',
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
        tests:            r.TestSelections ? r.TestSelections.split(',').map(t => t.trim()).filter(Boolean) : [],
        confidence:       r.OCRConfidence || 0,
        processedDate:    r.ProcessedDate || '',
        reviewStatus:     r.Title        || r.ReviewStatus || 'Pending',
        validationErrors: r.ValidationErrors || '',
        waterType:        r.WaterType    || '',
        _rowIndex:        r._id,
      }));

      // ── Recently approved (last 5 from Archived Intake) ──────────────────────
      const archivedItems = await listItems(LISTS.ARCHIVED_INTAKE, { top: 50 });

      const groupedByTs = {};
      archivedItems.forEach(r => {
        const ts = r.Timestamp || '';
        if (!ts) return;
        if (!groupedByTs[ts]) {
          groupedByTs[ts] = {
            ts,
            labIds:     [],
            coaTests:   [],
            customer:   r.ClientName || r.Customer || '',
            approvedBy: r.ReviewedBy || '',
          };
        }
        if (r.FullId)  groupedByTs[ts].labIds.push(r.FullId);
        if (r.CoaTest) groupedByTs[ts].coaTests.push(r.CoaTest);
      });

      const allSorted = Object.values(groupedByTs).sort((a, b) => {
        const baseA = (a.labIds[0]||'').match(/^\d{6}-\d{3}/)?.[0] || '';
        const baseB = (b.labIds[0]||'').match(/^\d{6}-\d{3}/)?.[0] || '';
        if (baseB > baseA) return 1;
        if (baseB < baseA) return -1;
        return 0;
      });

      const today = new Date().toLocaleDateString('en-CA', { timeZone: 'America/New_York' });
      const todayApproved = allSorted.slice(0, 5).map(g => ({
        ts:         g.ts,
        labIds:     g.labIds,
        tests:      g.coaTests,
        customer:   g.customer,
        approvedBy: g.approvedBy,
      }));
      const todayCount = allSorted.filter(g => g.ts.startsWith(today)).length;

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ pending, todayApproved, todayCount }),
      };

    } catch(e) {
      context.log('[get-scan-queue] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
