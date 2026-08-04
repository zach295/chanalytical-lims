const { app } = require('@azure/functions');
const { getToken, listItems } = require('../shared/graph');

app.http('clients-read', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const token   = await getToken();
      const siteId  = process.env.SP_SITE_ID;
      const GRAPH   = 'https://graph.microsoft.com/v1.0';

      const res = await fetch(
        `${GRAPH}/sites/${siteId}/lists/Clients/items?$expand=fields&$top=500&$orderby=fields/ClientName asc`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (!res.ok) throw new Error(`SP fetch failed: ${res.status}`);
      const data = await res.json();

      const clients = (data.value || []).map(item => {
        const f = item.fields || {};
        return {
          _id:              item.id,
          clientName:       f.ClientName       || '',
          clientCode:       f.ClientCode       || '',
          abbrev:           f.Abbrev           || '',
          mainContact:      f.Email            || '',   // Main Contact (first name)
          dbaName:          f.Phone            || '',   // DBA Name
          reportEmail:      f.Aliases          || '',   // Report Email Address
          billingEmail:     f.Notes            || '',   // Billing Email Address
          phone:            f.Active           || '',   // Phone #
          billingAddress:   f.BillingAddress   || '',
          billingPreference: f.BillingFrequency || '',
          frequency:        f.Frequency        || '',
          pricingCategory:  f.PricingCategory  || '',
          startDate:        f.StartDate        || '',
          status:           f.Status           || 'Active',
          radonLic:         f.RadonLic_x0023_  || '',
        };
      });

      return {
        status: 200,
        headers: { 'Cache-Control': 'no-store' },
        jsonBody: { clients },
      };
    } catch(e) {
      context.log('[clients-read] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    }
  }
});
