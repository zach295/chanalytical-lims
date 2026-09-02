/**
 * accession-status.js
 * GET  → { pending, reported }
 * POST → various actions (read-clients, today-approved, list-intake, mark-reported, mark-pending)
 */
const { app }                          = require('@azure/functions');
const { listItems, updateItem, LISTS } = require('../shared/graph');

function fmtName(name) {
  if (!name || !name.startsWith('Public-')) return name || '';
  const inner = name.slice(7).trim();
  const ci = inner.indexOf(',');
  if (ci > 0) {
    const last = inner.slice(0, ci).trim();
    const first = inner.slice(ci + 1).trim();
    return first ? `${first} ${last}` : last;
  }
  return inner;
}

app.http('accession-status', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {

      // ── GET ──────────────────────────────────────────────────────────────────
      if (request.method === 'GET') {
        const items = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 });
        const byBase = {};
        for (const r of items) {
          const fullId   = (r.field_1  || '').trim();
          const coaTest  = (r.field_2  || '').trim();
          const customer = (r.field_3  || '').trim();
          const status   = (r.field_14 || 'Pending').trim();
          if (!fullId) continue;
          const baseId = fullId.split(' ')[0].trim();
          if (!byBase[baseId]) {
            byBase[baseId] = {
              baseId, fullIds: [], tests: [],
              customer:   fmtName(customer),
              rawCustomer: customer,
              approvedBy: r.field_12 || '',
              timestamp:  r.Title    || '',
              status:     'Pending',
            };
          }
          byBase[baseId].fullIds.push(fullId);
          if (coaTest && !byBase[baseId].tests.includes(coaTest)) byBase[baseId].tests.push(coaTest);
          if (status && status !== 'Pending') byBase[baseId].status = status;
        }
        const all      = Object.values(byBase);
        const pending  = all.filter(r => r.status === 'Pending');
        const reported = all.filter(r => r.status !== 'Pending');
        return {
          status: 200,
          headers: { 'content-type': 'application/json', 'Cache-Control': 'no-store' },
          body: JSON.stringify({ pending, reported }),
        };
      }

      // ── POST ─────────────────────────────────────────────────────────────────
      if (request.method === 'POST') {
        const body   = await request.json().catch(() => ({}));
        const action = body.action || '';

        // overview-stats — returns raw records for the Overview tab
        if (action === 'overview-stats') {
          const allItems = await listItems(LISTS.ARCHIVED_INTAKE, { top: 2000 }).catch(() => []);
          const records  = allItems.map(r => ({
            labId:        (r.field_1  || '').trim(),
            testType:     (r.field_2  || '').trim(),
            customer:     fmtName(r.field_3 || ''),
            dateReceived: (r.field_6  || '').trim(),
            status:       (r.field_14 || 'Pending').trim(),
          })).filter(r => r.labId);
          return { status: 200, jsonBody: { records } };
        }

        // read-clients
        if (action === 'read-clients') {
          const { getToken } = require('../shared/graph');
          const token  = await getToken();
          const siteId = process.env.SP_SITE_ID;
          const GRAPH  = 'https://graph.microsoft.com/v1.0';
          const res    = await fetch(
            `${GRAPH}/sites/${siteId}/lists/Clients/items?$expand=fields&$top=500`,
            { headers: { Authorization: `Bearer ${token}` } }
          );
          if (!res.ok) return { status: 200, jsonBody: { clients: [] } };
          const data    = await res.json();
          const clients = (data.value || []).map(item => {
            const f = item.fields || {};
            return {
              _id:              item.id,
              clientName:       f.ClientName        || f.Title || '',
              clientCode:       f.ClientCode        || '',
              reportEmail:      f.Aliases           || '',
              billingEmail:     f.Notes             || '',
              phone:            f.Active            || '',
              dbaName:          f.Phone             || '',
              abbrev:           f.Abbrev            || '',
              billingAddress:   f.BillingAddress    || '',
              pricingCategory:  f.PricingCategory   || '',
              billingFrequency: f.BillingFrequency  || '',
              frequency:        f.Frequency         || '',
              startDate:        f.StartDate         || '',
              status:           f.Status            || '',
              radonLic:         f.RadonLic_x0023_   || '',
            };
          });
          return { status: 200, jsonBody: { clients } };
        }

        // today-approved
        if (action === 'today-approved') {
          const now      = new Date();
          const todayISO = now.toLocaleDateString('en-CA', { timeZone: 'America/New_York' }); // YYYY-MM-DD
          // Build MMDDYY prefix for lab ID matching (e.g. "080426" for 2026-08-04)
          const etParts  = new Intl.DateTimeFormat('en-US', {
            timeZone: 'America/New_York', month: '2-digit', day: '2-digit', year: '2-digit'
          }).formatToParts(now);
          const mm   = etParts.find(p=>p.type==='month')?.value || '';
          const dd   = etParts.find(p=>p.type==='day')?.value   || '';
          const yy   = etParts.find(p=>p.type==='year')?.value  || '';
          const labPrefix = `${mm}${dd}${yy}`; // e.g. "080426"

          const allI     = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 }).catch(() => []);
          const allToday = allI
            .filter(r => {
              const fullId = (r.field_1 || '').trim();
              const ts     = String(r.Title || '');
              return fullId.startsWith(labPrefix) || ts.startsWith(todayISO);
            })
            .map(r => ({
              fullId:     r.field_1  || '',
              coaTest:    r.field_2  || '',
              customer:   fmtName(r.field_3 || ''),
              approvedBy: r.field_12 || '',
              timestamp:  r.Title    || '',
            }))
            .sort((a, b) => b.fullId.localeCompare(a.fullId));

          // Build client breakdown from ALL today's items
          const clientCounts = {};
          allToday.forEach(a => {
            const c = a.customer || 'Unknown';
            clientCounts[c] = (clientCounts[c] || 0) + 1;
          });
          const clientBreakdown = Object.entries(clientCounts)
            .sort((a,b) => b[1]-a[1])
            .slice(0, 4)
            .map(([name, count]) => ({ name, count }));

          return { status: 200, jsonBody: {
            approved:         allToday.slice(0, 5), // top 5 for the table
            total:            allToday.length,
            clientBreakdown,                         // full breakdown for counter
          }};
        }

        // activity-log: returns all Activity Log entries from SharePoint
        if (action === 'activity-log') {
          const allItems = await listItems('Activity Log', { top: 2000 }).catch(e => {
            throw new Error('Activity Log fetch failed: ' + e.message);
          });
          const items = allItems.map(f => ({
            title:  f.Title        || '',
            client: f.Client       || '',
            type:   f.ActivityType || f.Type || '',
            notes:  f.Notes        || '',
            by:     f.By           || '',
            date:   f.LogDate      || '',
            time:   f.LogTime      || '',
            qty:    f.Quantity     || '',
          }));
          return { status: 200, jsonBody: { success: true, items } };
        }

        // clients-full: returns all client fields including Radon Lic # for state report
        if (action === 'clients-full') {
          const allClients = await listItems('Clients', { top: 500 }).catch(e => {
            throw new Error('Clients fetch failed: ' + e.message);
          });
          const clients = allClients.map(f => {
            const radonLic = f.RadonLic_x0023_
              || f.Radon_x0020_Lic_x0020__x0023_
              || f.RadonLic || f.Radon_x0020_Lic
              || f.RW_x0020_Lic || f.radonLic || '';
            const radonKeys = Object.keys(f).filter(k => /radon/i.test(k));
            return {
              clientCode:     f.ClientCode  || f.Title || '',
              clientName:     f.ClientName  || '',
              radonLic,
              radonKeys,
              billingZip:     (f.BillingAddress || '').match(/\d{5}/)?.[0] || '',
              billingAddress: f.BillingAddress || '',
            };
          });
          return { status: 200, jsonBody: { success: true, clients } };
        }

        // all-intake: returns all Archived Intake rows for the lab tab
        if (action === 'all-intake') {
          const allItems = await listItems(LISTS.ARCHIVED_INTAKE, { top: 2000 }).catch(() => []);
          const items = allItems.map(r => ({
            labId:        r.field_1 || '',
            services:     r.field_2 || '',
            clientName:   r.field_3 || '',
            dateDrawn:    r.field_4 || '',
            timeDrawn:    r.field_5 || '',
            dateReceived: r.field_6 || '',
            timeReceived: r.field_7 || '',
            location:     r.field_8 || '',
            city:         r.field_9 || '',
            state:        r.field_10 || '',
            zip:          r.field_11 || '',
            approvedBy:   r.field_12 || '',
            notes:        r.field_13 || '',
            status:       r.field_14 || '',
          }));
          return { status: 200, jsonBody: { success: true, items } };
        }

        // list-intake
        if (action === 'list-intake') {
          const searchId = body.baseId || '';
          if (!searchId) return { status: 400, jsonBody: { error: 'baseId required' } };
          const allItems = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 }).catch(() => []);
          const matches  = allItems.filter(r =>
            (r.field_1 || '').split(' ')[0].trim() === searchId
          ).map(r => ({
            _id: r._id, field_1: r.field_1||'', field_2: r.field_2||'',
            field_3: r.field_3||'', field_4: r.field_4||'', field_5: r.field_5||'',
            field_6: r.field_6||'', field_7: r.field_7||'', field_8: r.field_8||'',
            field_9: r.field_9||'', field_10: r.field_10||'', field_11: r.field_11||'',
            field_12: r.field_12||'', field_13: r.field_13||'', field_14: r.field_14||'',
            resolvedCoaTest: (r.field_2 || '').trim(),
          }));
          return { status: 200, jsonBody: { matches } };
        }

        // fix-field
        if (action === 'fix-field') {
          const { itemId, field, value } = body;
          await updateItem(LISTS.ARCHIVED_INTAKE, itemId, { [field]: value });
          return { status: 200, jsonBody: { success: true } };
        }

        // mark-reported / mark-pending
        const baseId = body.baseId || '';
        if (baseId && (action === 'mark-reported' || action === 'mark-pending')) {
          const newStatus = action === 'mark-reported' ? 'Reported' : 'Pending';
          const allItems  = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 }).catch(() => []);
          const matches   = allItems.filter(r => (r.field_1 || '').split(' ')[0].trim() === baseId);
          await Promise.all(matches.map(r =>
            updateItem(LISTS.ARCHIVED_INTAKE, r._id, { field_14: newStatus })
          ));
          return { status: 200, jsonBody: { success: true, baseId, status: newStatus } };
        }

        return { status: 400, jsonBody: { error: 'Unknown action: ' + action } };
      }

      return { status: 405, jsonBody: { error: 'Method not allowed' } };

    } catch(e) {
      context.log('[accession-status] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    }
  }
});
