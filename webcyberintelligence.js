const express = require('express');
const app = express();
const path = require('path');
const fs = require('fs');
const https = require('https');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const bodyParser = require('body-parser');
app.use(bodyParser.json());

//TODO Credenciales de Gajardo y client funcionando
const clientId = '35a4a79c-f0b7-499d-af22-d59364982ceb';
const clientSecret = 'qj_8Q~3yuRaR1s1g25OZ7PLNjErpLn4DF3.s1aKT';
const tenantId = 'ce6c3307-7b45-4cd2-a7ad-037c08909f1d';


//TODO parametros para crear una suscripcion
const webhookUrl = 'https://webhookprueba.ddns.net/sharepointCyberIntelligence';
const resource = `/drives/b!OPW09zb2_0yfKBe9foHcnK8u9Kmh-51Iq1uFWQ3AQqFNu9YMl6a2S44qkF0CLTOu/root`;
const currentDate = new Date();
const expirationDate =  new Date(currentDate.getTime() + 30 * 24 * 60 * 60 * 1000);
const subscriptionExpirationDateTime = expirationDate.toISOString();


//todo metodo de obtencion del token
const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
const client = Client.initWithMiddleware({
  authProvider: {
    getAccessToken: async () => {
      const token = await credential.getToken('https://graph.microsoft.com/.default');
      return token.token;
    },
  },
});

//todo metodo paar crear una suscripcion
async function createWebhookSubscription() {
    const subscription = {
        resource:resource,
        notificationUrl: webhookUrl,
        expirationDateTime: subscriptionExpirationDateTime,
        changeType: 'updated',
        clientState: 'secretClientValue',
    };

    try {
        const response = await client.api('/subscriptions').post(subscription);
        console.log('Webhook subscription created:', response);  
      } catch (error) {
        console.error('Error creating webhook subscription:', error);
     }
}

//todo metododpara eliminar una suscripcion
async function deletesuscription(){
    // Definir el subscriptionId de la suscripción a eliminar
  const subscriptionId = '6a58db34-8a5d-4afa-9fa4-c2e426a841dc'; // ID de la suscripción

  // Eliminar la suscripción
  client
    .api(`/subscriptions/${subscriptionId}`)
    .delete()
    .then(() => {
      console.log('Suscripción eliminada exitosamente');
    })
    .catch((error) => {
      console.error('Error al eliminar la suscripción:', error);
    });
}

//optimize: ruta del webhook en proceso para descragar todos los archivos
app.post('/sharepointCyberIntelligence', (req, res) => {
  const validationToken = req.query.validationToken;
  
  if (validationToken) { //validacion de la suscripcion
    res.status(200).send(validationToken);
    console.log('Validación de suscripción recibida:', validationToken);
  } else {
    // Notificación de actualización
    const notification = req.body;
    console.log('Notificación de actualización recibida:', notification);
    
    //si hay una notificacion pasa
    if(notification){
        const MainPath = 'sites/cybereye.sharepoint.com,f7b4f538-f636-4cff-9f28-17bd7e81dc9c,a9f42eaf-fba1-489d-ab5b-85590dc042a1/drives/b!OPW09zb2_0yfKBe9foHcnK8u9Kmh-51Iq1uFWQ3AQqFNu9YMl6a2S44qkF0CLTOu/root:/Ciberinteligencia/Clientes'
        const Customers = ['DEACERO'];    //'RCH-Bodega - CP','Grupo Real - CP'
        const informes = 'Informe Mensual';
        const fechaActual = new Date();
        const year = fechaActual.getFullYear();
        
        let Month;
        const mes_num = fechaActual.getMonth(); //+ 1; //Obtiene el mes 
        switch(mes_num){
          case 1:
            Month = 'Enero';
          break;
          case 2:
            Month = 'Febrero';
          break;
          case 3:
            Month = 'Marzo';
          break;
          case 4:
            Month = 'Abril';
          break;
          case 5:
            Month = 'Mayo';
          break;
          case 6:
            Month = 'Junio';
          break;
          case 7:
            Month = 'Julio';
          break;
          case 8:
            Month = 'Agosto';
          break;
          case 9:
            Month = 'Septiembre';
          break;
          case 10:
            Month = 'Octubre';
          break;
          case 11:
            Month = 'Noviembre';
          break;
          case 12:
            Month = 'Diciembre';
          break;
        }
        
    
        
        //TODO DESCARGA DE ARCHIVOS
        
        async function iterateArray() {
              for (const element of Customers){
                      const resListado = await client.api(`${MainPath}/${element}/${informes}/${year}/${Month}:/children`).get();
                      const ret = resListado.value;
                      for (let x of ret) {
                        if(x.name.split('.').pop() == 'pdf'){
                          (async () => {
                            
                            const response = await client.api(`${MainPath}/${element}/${informes}/${year}/${Month}/${x.name}`).get();
                            const downloadUrl = response['@microsoft.graph.downloadUrl'];
                            //canjeo del url por medio del client y le decimos que nos repondas un tipo arraybuffer
                            const responseArchivo = await client.api(downloadUrl).responseType('arraybuffer').get();
                            //nos permite crear un nuevo buffer en el que contendremos una cadena de caracteres
                            //en pocas palabras es el que convierte el buffer obtenido a strings 
                            const archivoBuffer = Buffer.from(responseArchivo);
        
                            //guardamos el archivos  que tiene el archivo convertido 
                            fs.writeFileSync(path.resolve(__dirname, `./descargado/${x.name}`), archivoBuffer);
                          })();
                        } 
                    }  
                    console.log('Arcchivo Descargado Así nomas alv')
                  
              }
          }
      iterateArray();
    }
    
    res.status(200).send('Notificación recibida');
  }
    
});


/*le diecimos a el servidor que todo lo que esta dentro de la carpeta public sera accesible 
por le navegador*/ 
app.use(express.static(path.join(__dirname,'public')));



https.createServer({
  cert: fs.readFileSync('certificate.crt'),
  key: fs.readFileSync('private.key'),

}, app).listen(3333, () => {
  console.log('Servidor iniciado en el puerto 4000');
  /**
   * //todo estos metodos estan comentados ya que uno crea una suscripcion despues de que se crea la suscripcion ya no es necesario porque creamas 
  */
  createWebhookSubscription();
  //deletesuscription();
});