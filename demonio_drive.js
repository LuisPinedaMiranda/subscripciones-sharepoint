const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');

const clientId = '35a4a79c-f0b7-499d-af22-d59364982ceb';
const clientSecret = 'qj_8Q~3yuRaR1s1g25OZ7PLNjErpLn4DF3.s1aKT';
const tenantId = 'ce6c3307-7b45-4cd2-a7ad-037c08909f1d';

const creds = new ClientSecretCredential(tenantId, clientId, clientSecret);
const client = Client.initWithMiddleware({
  authProvider: async (done) => {
    const token = await creds.getToken(['https://graph.microsoft.com/.default']);
    done(null, token.accessToken);
  }
});

async function listarDocumentos() {
    try {
      const directorioCompartidoId = 'ID-o-Ruta-DirectorioCompartido';
      const response = await client.api(`/sites/cybereye.sharepoint.com,ba55d390-b892-4438-abcc-dfed349d8958/drive/root:/SOC/Clientes/DEACERO/Reporte Mensual/Abril:/children`).get();
      console.log(response.value);
    } catch (error) {
      console.log(error);
    }
  }
  
  listarDocumentos();
  