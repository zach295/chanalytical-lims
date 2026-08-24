const { app } = require('@azure/functions');
const { getToken, createItem, updateItem } = require('../shared/graph');

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

async function logActivity(action, clientName, details, by) {
  try {
    const now     = new Date();
    const logDate = now.toLocaleDateString('en-US',{ timeZone:'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });
    const logTime = now.toLocaleTimeString('en-US',{ timeZone:'America/New_York', hour:'2-digit', minute:'2-digit', hour12:false });
    await createItem('Activity Log', {
      Title: `${logDate} ${clientName}`, Client: clientName,
      ActivityType: action, Notes: details.slice(0, 3000),
      By: by || 'Admin', LogDate: logDate, LogTime: logTime, Quantity: 0,
    });
  } catch(e) {}
}

app.http('clients-write', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const body   = await request.json();
      const { action, itemId, client, updatedBy } = body;
      if (!client?.clientName) return { status: 400, jsonBody: { error: 'clientName required' } };
      const token  = await getToken();
      const fields = toSpFields(client);

      if (action === 'add') {
        const result = await createItem(LISTS.CLIENTS, fields);
        const details = `New client added to Clients list.\nCode: ${client.clientCode || '—'} | Pricing: ${client.pricingCategory || '—'} | Status: ${client.status || 'Active'} | Report Email: ${client.reportEmail || '—'} | Phone: ${client.phone || '—'} | Billing Address: ${client.billingAddress || '—'}`;
        await logActivity('Client Added', client.clientName, details, updatedBy);
        return { status: 200, jsonBody: { success: true, id: result?.id } };
      }

      if (action === 'update') {
        if (!itemId) return { status: 400, jsonBody: { error: 'itemId required for update' } };
        await updateItem(LISTS.CLIENTS, itemId, fields);
        const changed = Object.entries(client).filter(([k,v]) => v !== undefined && v !== '').map(([k,v]) => `${k}: ${v}`).join(' | ');
        const details = `Clients list updated.\nFields changed: ${changed}`;
        await logActivity('Client Updated', client.clientName, details, updatedBy);
        return { status: 200, jsonBody: { success: true } };
      }

      if (action === 'deactivate' || action === 'activate') {
        if (!itemId) return { status: 400, jsonBody: { error: 'itemId required' } };
        const newStatus = action === 'activate' ? 'Active' : 'Inactive';
        await updateItem(LISTS.CLIENTS, itemId, { Status: newStatus });
        const details = `Status changed to ${newStatus} in Clients list.`;
        await logActivity(`Client ${newStatus}`, client.clientName, details, updatedBy);
        return { status: 200, jsonBody: { success: true } };
      }

      return { status: 400, jsonBody: { error: `Unknown action: ${action}` } };
    } catch(e) {
      context.log('[clients-write] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    }
  }
});
