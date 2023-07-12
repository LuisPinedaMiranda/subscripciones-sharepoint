const fs = require('fs');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const path = require('path')

//Credenciales de Gajardo
const clientId = '35a4a79c-f0b7-499d-af22-d59364982ceb';
const clientSecret = 'qj_8Q~3yuRaR1s1g25OZ7PLNjErpLn4DF3.s1aKT';
const tenantId = 'ce6c3307-7b45-4cd2-a7ad-037c08909f1d';


//Credenciales de mi cuenta de zacatepec
/*const clientId = '06a5db4f-068d-40ed-9787-9d52f177ba26';
const clientSecret = 'CrU8Q~rcjTIHj0WuPMyERAh0FA3BYJeJfIDbUcZL';
const tenantId = '261809a4-d6b4-48b4-a568-f07df069157c';*/


const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
const client = Client.initWithMiddleware({
  authProvider: {
    getAccessToken: async () => {
      const token = await credential.getToken('https://graph.microsoft.com/.default');
      return token.token;
    },
  },
});

const downloadFile = async (userId, fileId, name) => {
  try {
    const response = await client.api(`/users/${userId}/drive/items/${fileId}`).get();
    const downloadUrl = response['@microsoft.graph.downloadUrl'];


    //canjeo del url por medio del client y le decimos que nos repondas un tipo arraybuffer
    const responseArchivo = await client.api(downloadUrl).responseType('arraybuffer').get();
    //nos permite crear un nuevo buffer en el que contendremos una cadena de caracteres
    //en pocas palabras es el que convierte el buffer obtenido a strings 
    const archivoBuffer = Buffer.from(responseArchivo);

    //guardamos el archivos  que tiene el archivo convertido 
    fs.writeFileSync(path.resolve(__dirname, `./descargado/${name}.docx`), archivoBuffer);

  } catch (error) {
    console.log('Error al descargar el archivo:', error);
  }
};

const fileId = '01KMZ6RXXUH5GTHHROQRA3E6WYRJGLEQKP'; // Reemplaza con el ID real del archivo que deseas descargar
const name = `luis_olvera`; // Ruta y nombre del archivo de destino
const userId = '3e273bff-66d9-430a-a19b-25aaf48597d8';





downloadFile(userId, fileId,name)
  .then(() => {
    console.log('Archivo descargado exitosamente.');
  })
  .catch((error) => {
    console.log('Error al descargar el archivo:', error);
  });
















/*const creds = new ClientSecretCredential(tenantId, clientId, clientSecret);
const client = Client.initWithMiddleware({
  authProvider: {
    getAccessToken: async () => {
      const token = await creds.getToken(['https://graph.microsoft.com/.default']);
      return token.token;
    }
  }
});*/








// Función para descargar el archivo
















/*async function listarDocumentos() {
  try {
    const response = await client.api('/sites/cybereye.sharepoint.com,ba55d390-b892-4438-abcc-dfed349d8958/drive/root:/SOC/Clientes/DEACERO/Reporte Mensual/Abril:/children').get();
    console.log(response.value);
  } catch (error) {
    console.log(error);
  }
}

listarDocumentos();*/
