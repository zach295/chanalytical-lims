const { app } = require('@azure/functions');
const { getToken, createItem, updateItem, listItems } = require('../shared/graph');
const { writeActivityLog } = require('../shared/audit');

const LISTS = { CLIENTS: 'Clients' };

function toSpFields(c) {
  return {
    ClientName:       c.clientName      || '',
    ClientCode:       c.clientCode      || '',
    Abbrev:           c.abbrev          || '',
    Email:            c.mainContact     || '',
    Phone:            c.dbaName         || '',
    Aliases:          c.reportEmail     || '',
    Notes:            c.billingEmail    || '',
    Active:           c.phone           || '',
    BillingAddress:   c.billingAddress  || '',
    BillingFrequency: c.billingPreference || '',
    Frequency:        c.frequency       || '',
    PricingCategory:  c.pricingCategory || '',
    StartDate:        c.startDate       || '',
    Status:           c.status          || 'Active',
    RadonLic_x0023_:  c.radonLic        || '',
  };
}

function diffFields(oldFields, newFields) {
  return Object.entries(newFields)
    .filter(([k, v]) => String(oldFields?.[k] ?? '') !== String(v ?? ''))
    .map(([k, v]) => `${k}: "${oldFields?.[k] ?? ''}" → "${v ?? ''}"`);
}

app.http('clients-write', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const body = await request.json();
      const { action, itemId, client, updatedBy } = body;
      if (!client?.clientName) return { status: 400, jsonBody: { error: 'clientName required' } };
      await getToken(); // preserve existing auth/connection check
      const normalizedClient = { ...client };
      // New manually-added public clients follow the same Last, First naming
      // convention as public clients created by Review Queue.
      if (action === 'add' && String(client.pricingCategory || '').trim().toLowerCase() === 'public pricing') {
        const raw = String(client.clientName || '').trim().replace(/^Public-/i, '').trim();
        const parts = raw.split(/\s+/).filter(Boolean);
        let last = '', first = '';
        if (raw.includes(',')) {
          const comma = raw.indexOf(',');
          last = raw.slice(0, comma).trim();
          first = raw.slice(comma + 1).trim();
        } else if (parts.length > 1) {
          last = parts.pop();
          first = parts.join(' ');
        } else {
          last = raw;
        }
        normalizedClient.clientName = first ? `${last}, ${first}` : last;
        normalizedClient.dbaName = first ? `${first} ${last}` : last;
        normalizedClient.abbrev = 'PUBLIC';
        normalizedClient.billingPreference = 'Pre-Pay';
      }
      const fields = toSpFields(normalizedClient);
      const actor = updatedBy || 'Admin';

      if (action === 'add') {
        const result = await createItem(LISTS.CLIENTS, fields);
        const details = `New client added. Code: ${client.clientCode || '—'} | Pricing: ${client.pricingCategory || '—'} | Status: ${client.status || 'Active'} | Report Email: ${client.reportEmail || '—'} | Phone: ${client.phone || '—'} | Billing Address: ${client.billingAddress || '—'}`;
        const audit = await writeActivityLog({ labId: normalizedClient.clientName, type: 'Client Added', notes: details, by: actor, context });
        return { status: 200, jsonBody: { success: true, id: result?.id, auditWarning: audit.success ? null : audit.error } };
      }

      if (action === 'update') {
        if (!itemId) return { status: 400, jsonBody: { error: 'itemId required for update' } };
        const old = (await listItems(LISTS.CLIENTS, { top: 1000 }).catch(() => []))
          .find(r => String(r._id) === String(itemId)) || {};
        const changes = diffFields(old, fields);
        await updateItem(LISTS.CLIENTS, itemId, fields);
        const audit = await writeActivityLog({
          labId: client.clientName,
          type: 'Client Updated',
          notes: changes.length ? changes.join(' | ') : 'Client update saved with no material field change.',
          by: actor,
          context,
        });
        return { status: 200, jsonBody: { success: true, auditWarning: audit.success ? null : audit.error } };
      }

      if (action === 'deactivate' || action === 'activate') {
        if (!itemId) return { status: 400, jsonBody: { error: 'itemId required' } };
        const newStatus = action === 'activate' ? 'Active' : 'Inactive';
        const old = (await listItems(LISTS.CLIENTS, { top: 1000 }).catch(() => []))
          .find(r => String(r._id) === String(itemId)) || {};
        await updateItem(LISTS.CLIENTS, itemId, { Status: newStatus });
        const audit = await writeActivityLog({
          labId: client.clientName,
          type: `Client ${newStatus}`,
          notes: `Status: "${old.Status || ''}" → "${newStatus}"`,
          by: actor,
          context,
        });
        return { status: 200, jsonBody: { success: true, auditWarning: audit.success ? null : audit.error } };
      }

      return { status: 400, jsonBody: { error: `Unknown action: ${action}` } };
    } catch(e) {
      context.log('[clients-write] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    }
  }
});
