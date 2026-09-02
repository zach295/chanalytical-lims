const { app } = require('@azure/functions');

app.http('health', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    context.log('Health function reached');
    return {
      status: 200,
      jsonBody: { ok: true, service: 'chanalytical-api' }
    };
  }
});
