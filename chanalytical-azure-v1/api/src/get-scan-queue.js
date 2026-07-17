const { app } = require('@azure/functions');
const { listItems, findItem, LISTS } = require('../shared/graph');

app.http('get-scan-queue', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      // ── Pending scans from Review Queue ──────────────────────────────────────
      const queueItems = await listItems(LISTS.REVIEW_QUEUE, {
        filter: "fields/ReviewStatus ne 'Approved' and fields/ReviewStatus ne 'Discarded'",
        orderby: 'fields/ProcessedDate desc',
        top: 100,
      });

      const pending = queueItems.map(r => ({
        fileId:         r.FileId         || '',
        barcodeId:      r.BarcodeId      || '',
        customer:       r.ClientName     || '',
        email:          r.Email          || '',
        dateDrawn:      r.SampleDate     || '',
        timeDrawn:      r.SampleTime     || '',
        receivedDate:   r.ReceivedDate   || '',
        receivedTime:   r.ReceivedTime   || '',
        location:       r.Address        || '',
        city:           r.City           || '',
        state:          r.State          || 'ME',
        zip:            r.Zip            || '',
        tests:          r.TestSelections ? r.TestSelections.split(',').map(t => t.trim()).filter(Boolean) : [],
        confidence:     r.OcrConfidence  || 0,
        processedDate:  r.ProcessedDate  || '',
        reviewStatus:   r.ReviewStatus   || 'Pending',
        validationErrors: r.ValidationErrors || '',
        _rowIndex:      r._id,
      }));

      // ── Recently approved (last 5 kits from Archived Intake) ─────────────────
      const archivedItems = await listItems(LISTS.ARCHIVED_INTAKE, {
        orderby: 'fields/Created desc',
        top: 50,
      });

      // Group by timestamp to combine kits approved together
      const groupedByTs = {};
      archivedItems.forEach(r => {
        const ts = r.Timestamp || '';
        if (!ts) return;
        if (!groupedByTs[ts]) {
          groupedByTs[ts] = {
            ts,
            labIds:    [],
            coaTests:  [],
            customer:  r.Customer   || '',
            approvedBy: r.ReviewedBy || '',
          };
        }
        if (r.FullId)  groupedByTs[ts].labIds.push(r.FullId);
        if (r.CoaTest) groupedByTs[ts].coaTests.push(r.CoaTest);
      });

      const allSorted = Object.values(groupedByTs)
        .sort((a, b) => {
          const seq = ids => Math.max(...ids.map(id => {
            const m = id.match(/-(\d{3})/);
            return m ? parseInt(m[1]) : 0;
          }));
          const baseA = (a.labIds[0]||'').match(/^\d{6}-\d{3}/)?.[0]||'';
          const baseB = (b.labIds[0]||'').match(/^\d{6}-\d{3}/)?.[0]||'';
          if (baseB > baseA) return 1;
          if (baseB < baseA) return -1;
          return seq(b.labIds) - seq(a.labIds);
        });

      const today = new Date().toLocaleDateString('en-CA', { timeZone: 'America/New_York' });
      const todayApproved = allSorted.slice(0, 5).map(g => ({
        ts:        g.ts,
        labIds:    g.labIds,
        tests:     g.coaTests,
        customer:  g.customer,
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
