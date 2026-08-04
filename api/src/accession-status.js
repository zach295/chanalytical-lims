/**
 * accession-status.js — Azure version
 * Reads Archived Intake (field_X mapping) to list pending/reported Lab IDs.
 *
 * GET  → { pending: [...], reported: [...] }
 * POST { action:'mark-reported', baseId } → sets field_14 to 'Reported'
 * POST { action:'mark-pending',  baseId } → sets field_14 to 'Pending'
 */
const { app } = require('@azure/functions');
const { listItems, updateItem, getToken, LISTS } = require('../shared/graph');
const GRAPH = 'https://graph.microsoft.com/v1.0';

// Format public client name: "Public-Chandler, Zach" → "Zach Chandler"
function formatCustomerName(name) {
  if (!name) return '';
  if (!name.startsWith('Public-')) return name;
  // Strip "Public-" prefix
  const inner = name.slice('Public-'.length).trim();
  // Check for "Last, First" format
  const commaIdx = inner.indexOf(',');
  if (commaIdx > 0) {
    const last  = inner.slice(0, commaIdx).trim();
    const first = inner.slice(commaIdx + 1).trim();
    return first ? `${first} ${last}` : last;
  }
  // No comma — just return the name without "Public-"
  return inner;
}

// v202608040008 — force instance reload
// Archived Intake field mapping:
// Title=timestamp, field_1=fullId, field_2=coaTest, field_3=clientName,
// field_4=dateDrawn, field_5=timeDrawn, field_6=receivedDate, field_7=receivedTime,
// field_8=address, field_9=city, field_10=state, field_11=zip,
// field_12=approvedBy, field_13=notes, field_14=status

app.http('accession-status', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      if (request.method === 'GET') {
        const items = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 });

        // Group by base Lab ID
        const byBase = {};
        for (const r of items) {
          const fullId   = (r.field_1  || '').trim();
          const coaTest  = (r.field_2 || r.Test || r['Test0'] || '').trim();
          const customer = (r.field_3  || '').trim();
          const status   = (r.field_14 || 'Pending').trim();
          if (!fullId) continue;

          const baseId = fullId.split(' ')[0].trim();
          if (!byBase[baseId]) {
            byBase[baseId] = {
              baseId,
              fullIds:      [],
              tests:        [],
              customer: formatCustomerName(customer),
              location:     r.field_8  || '',
              city:         r.field_9  || '',
              state:        r.field_10 || 'ME',
              zip:          r.field_11 || '',
              dateDrawn:    r.field_4  || '',
              timeDrawn:    r.field_5  || '',
              dateReceived: r.field_6  || '',
              timeReceived: r.field_7  || '',
              approvedBy:   r.field_12 || '',
              email:        '',
              status:       'Pending',
              timestamp:    r.Title    || '',
              _ids:         [],
            };
          }
          byBase[baseId]._ids.push(r._id);
          if (fullId  && !byBase[baseId].fullIds.includes(fullId))  byBase[baseId].fullIds.push(fullId);
          if (coaTest && !byBase[baseId].tests.includes(coaTest))   byBase[baseId].tests.push(coaTest);
          if (status === 'Sent' || status === 'Reported')            byBase[baseId].status = 'Sent';
          if (!byBase[baseId].customer && customer)                  byBase[baseId].customer = customer;
        }

        const all      = Object.values(byBase);
        const pending  = all.filter(r => r.status !== 'Sent' && r.status !== 'Reported');
        const reported = all.filter(r => r.status === 'Sent' || r.status === 'Reported');

        // Fetch client emails — direct Graph API call using same pattern as clients-read
        try {
          const token2 = await getToken();
          const siteId2 = process.env.SP_SITE_ID;
          const cRes = await fetch(
            `${GRAPH}/sites/${siteId2}/lists/Clients/items?$expand=fields&$top=500`,
            { headers: { Authorization: `Bearer ${token2}` } }
          );
          if (!cRes.ok) throw new Error(`Clients fetch: ${cRes.status}`);
          const cJson = await cRes.json();
          const emailMap = {};
          for (const item of (cJson.value || [])) {
            const f = item.fields || {};
            const name  = String(f.Title || f.ClientName || '').trim();
            const email = String(f.Email || '').trim();
            if (!name || !email) continue;
            // Index by raw name and formatted name
            emailMap[name.toLowerCase()] = email;
            const fmt = formatCustomerName(name).toLowerCase();
            if (fmt !== name.toLowerCase()) emailMap[fmt] = email;
          }
          // Build raw customer map from intake items
          const rawByBase = {};
          for (const r of intakeItems) {
            const bid = (r.field_1 || '').split(' ')[0].trim();
            if (bid && r.field_3 && !rawByBase[bid]) rawByBase[bid] = r.field_3.trim();
          }
          for (const entry of pending) {
            const raw  = (rawByBase[entry.baseId] || '').toLowerCase();
            const disp = (entry.customer || '').toLowerCase();
            entry.email = emailMap[raw] || emailMap[disp] || '';
          }
        } catch(e) { console.warn('[accession-status] email lookup failed:', e.message); }

        return {
          status: 200,
          headers: {
            'content-type': 'application/json',
            'Cache-Control': 'no-store, no-cache, must-revalidate, max-age=0',
            'Pragma': 'no-cache',
          },
          body: JSON.stringify({ pending, reported }),
        };
      }

      if (request.method === 'POST') {
        const body2 = await request.json().catch(() => ({}));
        const { action, baseId } = body2;

        // Read full Clients list with new field names
        if (action === 'read-clients') {
          const tkn    = await getToken();
          const siteId = process.env.SP_SITE_ID;
          const GRAPH  = 'https://graph.microsoft.com/v1.0';
          const cRes   = await fetch(
            `${GRAPH}/sites/${siteId}/lists/Clients/items?$expand=fields&$top=500`,
            { headers: { Authorization: `Bearer ${tkn}` } }
          );
          if (!cRes.ok) return { status: 500, jsonBody: { error: `Clients fetch failed: ${cRes.status}` } };
          const cData = await cRes.json();
          const clients = (cData.value || []).map(item => {
            const f = item.fields || {};
            return {
              _id:              item.id,
              clientName:       f.ClientName       || '',
              clientCode:       f.ClientCode       || '',
              abbrev:           f.Abbrev           || '',
              mainContact:      f.Email            || '',
              dbaName:          f.Phone            || '',
              reportEmail:      f.Aliases          || '',
              billingEmail:     f.Notes            || '',
              phone:            f.Active           || '',
              billingAddress:   f.BillingAddress   || '',
              billingPreference: f.BillingFrequency || '',
              frequency:        f.Frequency        || '',
              pricingCategory:  f.PricingCategory  || '',
              startDate:        f.StartDate        || '',
              status:           f.Status           || 'Active',
              radonLic:         f.RadonLic_x0023_  || '',
            };
          });
          return { status: 200, jsonBody: { clients } };
        }

        // Direct field patch action
        if (action === 'fix-field') {
          const { itemId, field, value } = body2;
          await updateItem(LISTS.ARCHIVED_INTAKE, itemId, { [field]: value });
          return { status: 200, jsonBody: { success: true, itemId, field, value } };
        }

        // Mirror the GET handler exactly — proves what the deployed GET code does
        if (action === 'debug-get') {
          const { baseId: searchId } = body2;
          const items = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 });
          const results = [];
          for (const r of items) {
            const fullId  = (r.field_1 || '').trim();
            const coaTest = (r.field_2 || r.Test || r['Test0'] || '').trim();
            const baseId2 = fullId.split(' ')[0].trim();
            if (baseId2 !== searchId) continue;
            results.push({ _id: r._id, fullId, field_2: r.field_2, Test: r.Test, coaTest });
          }
          return { status: 200, jsonBody: { results, itemsTotal: items.length } };
        }

        // List ALL Archived Intake rows matching a baseId
        if (action === 'list-intake') {
          const { baseId: searchId } = body2;
          const allItems = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 });
          const matches  = allItems.filter(r => {
            const fid = (r.field_1 || '').split(' ')[0].trim();
            return fid === searchId;
          }).map(r => ({
            _id:          r._id,
            field_1:      r.field_1  || '',
            field_2:      r.field_2  || '',
            field_3:      r.field_3  || '',
            field_4:      r.field_4  || '',
            field_5:      r.field_5  || '',
            field_6:      r.field_6  || '',
            field_7:      r.field_7  || '',
            field_8:      r.field_8  || '',
            field_9:      r.field_9  || '',
            field_10:     r.field_10 || '',
            field_11:     r.field_11 || '',
            field_12:     r.field_12 || '',
            field_13:     r.field_13 || '',
            field_14:     r.field_14 || '',
            resolvedCoaTest: (r.field_2 || r.Test || '').trim(),
          }));
          return { status: 200, jsonBody: { matches, count: matches.length } };
        }

        // Read raw fields from a specific item — for debugging
        if (action === 'read-raw') {
          const { itemId } = body2;
          const siteId = process.env.SP_SITE_ID;
          const tkn    = await getToken();
          const lists  = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$select=id,displayName`,
            { headers: { Authorization: `Bearer ${tkn}` } }).then(r=>r.json());
          const lst    = (lists.value||[]).find(l => l.displayName === 'Archived Intake');
          if (!lst) return { status: 404, jsonBody: { error: 'list not found' } };
          const item   = await fetch(
            `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${lst.id}/items/${itemId}/fields`,
            { headers: { Authorization: `Bearer ${tkn}` } }).then(r=>r.json());
          return { status: 200, jsonBody: { fields: item } };
        }
        if (!baseId) return { status: 400, body: JSON.stringify({ error: 'baseId required' }) };

        const newStatus = action === 'mark-reported' ? 'Reported' : 'Pending';

        // Find all Archived Intake rows matching this base ID
        const items   = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 });
        const matches = items.filter(r => {
          const fullId = (r.field_1 || '').trim();
          return fullId && fullId.replace(/\s+\S+$/, '').trim() === baseId;
        });

        if (!matches.length) {
          return { status: 404, body: JSON.stringify({ error: `No rows found for ${baseId}` }) };
        }

        await Promise.all(matches.map(r =>
          updateItem(LISTS.ARCHIVED_INTAKE, r._id, { field_14: newStatus })
        ));

        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ success: true, baseId, status: newStatus, rowsUpdated: matches.length }),
        };
      }

      return { status: 405, body: 'Method Not Allowed' };

    } catch(e) {
      context.log('[accession-status] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
