const { app } = require('@azure/functions');
const { getToken, createItem, updateItem } = require('../shared/graph');

const LISTS = { CLIENTS: 'Clients' };

function toSpFields(c) {
  // Map dashboard field names to SharePoint internal field names
  return {
    ClientName:       c.clientName      || '',
    ClientCode:       c.clientCode      || '',
    Abbrev:           c.abbrev          || '',
    Email:            c.mainContact     || '',  // Main Contact (first name)
    Phone:            c.dbaName         || '',  // DBA Name
    Aliases:          c.reportEmail     || '',  // Report Email Address
    Notes:            c.billingEmail    || '',  // Billing Email Address
    Active:           c.phone           || '',  // Phone #
    BillingAddress:   c.billingAddress  || '',
    BillingFrequency: c.billingPreference || '',
    Frequency:        c.frequency       || '',
    PricingCategory:  c.pricingCategory || '',
    StartDate:        c.startDate       || '',
    Status:           c.status          || 'Active',
    RadonLic_x0023_:  c.radonLic        || '',
  };
}

app.http('clients-write', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const body   = await request.json();
      const { action, itemId, client } = body;

      if (!client?.clientName) return { status: 400, jsonBody: { error: 'clientName required' } };

      const token  = await getToken();
      const fields = toSpFields(client);

      if (action === 'add') {
        const result = await createItem(LISTS.CLIENTS, fields);
        return { status: 200, jsonBody: { success: true, id: result?.id } };
      }

      if (action === 'update') {
        if (!itemId) return { status: 400, jsonBody: { error: 'itemId required for update' } };
        await updateItem(LISTS.CLIENTS, itemId, fields);
        return { status: 200, jsonBody: { success: true } };
      }

      if (action === 'deactivate' || action === 'activate') {
        if (!itemId) return { status: 400, jsonBody: { error: 'itemId required' } };
        await updateItem(LISTS.CLIENTS, itemId, { Status: action === 'activate' ? 'Active' : 'Inactive' });
        return { status: 200, jsonBody: { success: true } };
      }

      return { status: 400, jsonBody: { error: `Unknown action: ${action}` } };
    } catch(e) {
      context.log('[clients-write] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    }
  }
});
