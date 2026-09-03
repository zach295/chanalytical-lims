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
      const fields = toSpFields(client);
      const actor = updatedBy || 'Admin';

      if (action === 'add') {
        const result = await createItem(LISTS.CLIENTS, fields);
        const details = `New client added. Code: ${client.clientCode || '—'} | Pricing: ${client.pricingCategory || '—'} | Status: ${client.status || 'Active'} | Report Email: ${client.reportEmail || '—'} | Phone: ${client.phone || '—'} | Billing Address: ${client.billingAddress || '—'}`;
        const audit = await writeActivityLog({ labId: client.clientName, type: 'Client Added', notes: details, by: actor, context });
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
