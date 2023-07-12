const fs = require('fs');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const path = require('path');
const { get } = require('http');

//TODO Credenciales de Gajardo y client funcionando
const clientId = '35a4a79c-f0b7-499d-af22-d59364982ceb';
const clientSecret = 'qj_8Q~3yuRaR1s1g25OZ7PLNjErpLn4DF3.s1aKT';
const tenantId = 'ce6c3307-7b45-4cd2-a7ad-037c08909f1d';

const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

const client = Client.initWithMiddleware({
  authProvider: {
    getAccessToken: async () => {
      const token = await credential.getToken('https://graph.microsoft.com/.default');
      console.log(token);
      return token.token;
    },
  },
});
//console.log(client);


//OPTIMIZE checando el pentest de los clientes
const ruta =`sites/cybereye.sharepoint.com,ba55d390-b892-4438-abcc-dfed349d8958/drive/root:/SOC/Clientes`

const Clientes = ['PAVISA', 'RCH-Bodega_RCH', 'ConsuBanco','DEACERO','Exitus', 'GrupoReal',
                   'Interproteccion', 'Magnet', 'Qualitas CR', 'Qualitas MX', 'TME'];
        
    /*for(let c of Clientes){
        console.log(c);
    }*/

//const clien=Clientes[0]

  //Función para obtener mes actual.  
  /* const getMnt = () =>{
    var fechaActual = new Date();
    var mesActual = fechaActual.getMonth();
    mesActual++; // Aumentar en 1 para que enero sea 1 en lugar de 0
    return mesActual;
   }
   
   //Funcion para obtener año actual
   const getYr = () => {
    var fechaActual = new Date();
    var añoActual = fechaActual.getFullYear();
    return añoActual;
   }

   const listado = async (año,mes) => {
    const resListado = await client.api(`${ruta}/${clien}/Mensual/${año}/${mes}:/children`).get();
    console.log(resListado);
    //return resListado.value;
  };
   
    

  if(client){ 
    const mes_num = 5;
    const año = getYr();
    let mes;

    switch(mes_num){
      case 1:
      break;
      case 2:
      break;
      case 3:
      break;
      case 4:
      break;
      case 5:
       mes = 'Mayo';
       listado(año,mes);
      break;
      case 6:
        mes = 'Junio';
      break;
      case 7:
      break;


    }
     
  }*/


 







/*-------------------------------------------------------------*/ 

//TODO DESCARGA DE ARCHIVOS

const listado = async () => {
    const resListado = await client.api(`sites/cybereye.sharepoint.com,ba55d390-b892-4438-abcc-dfed349d8958/drive/root:/SOC/Cyberpeace/Clientes/DeAcero - CP/Informe Mensual/2023/Mayo:/children`).get();
    return resListado.value;
};

//console.log(client);

const obtenerIds = async () => {
    const ids = await listado();

    for (let x of ids) {
        if(x.name.split('.').pop() == 'pdf'){
            
            const response = await client.api(`sites/cybereye.sharepoint.com,ba55d390-b892-4438-abcc-dfed349d8958/drive/root:/SOC/Cyberpeace/Clientes/DeAcero - CP/Informe Mensual/2023/Mayo/${x.name}`).get();
            const downloadUrl = response['@microsoft.graph.downloadUrl'];
            //canjeo del url por medio del client y le decimos que nos repondas un tipo arraybuffer
            const responseArchivo = await client.api(downloadUrl).responseType('arraybuffer').get();
            //nos permite crear un nuevo buffer en el que contendremos una cadena de caracteres
            //en pocas palabras es el que convierte el buffer obtenido a strings 
            const archivoBuffer = Buffer.from(responseArchivo);

            //guardamos el archivos  que tiene el archivo convertido 
            fs.writeFileSync(path.resolve(__dirname, `./descargado/${x.name}`), archivoBuffer);
            
        }

      }  
      console.log('Archivos Descargados');
};


obtenerIds();
