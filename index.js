process.title = "WP GRUPO ORIENTAL";
const qrcode = require('qrcode-terminal');
const fs = require('fs');
const { Client } = require('whatsapp-web.js');
const watsonResponse = require('./watsonResponse.js')

const { MessageMedia } = require('whatsapp-web.js')
var http = require('http');
var https = require('https');
const sqlController = require('./sqlcontroller');
require('tls').DEFAULT_MIN_VERSION = 'TLSv1';
//const client = new Client();
var fetch = require('node-fetch');
const SESSION_FILE_PATH = 'session.json';
// Load the session data if it has been previously saved
let sessionData;
if (fs.existsSync(SESSION_FILE_PATH)) {
    let rawdata = fs.readFileSync('session.json');
    let wTcODE = JSON.parse(rawdata);
    sessionData = wTcODE;
}
// Use the saved values
const client = new Client({
    session: sessionData
});

// Save session values to the file upon successful auth
client.on('authenticated', (session) => {
    sessionData = session;
    fs.writeFile(SESSION_FILE_PATH, JSON.stringify(session), function (err) {
        if (err) {
            console.error(err);
        }
    });
});


client.on('qr', qr => {
    qrcode.generate(qr, { small: true });
});
client.on('ready', () => {
    console.log('Client is ready!' + '***' + client.info.me._serialized + '****');
    console.log('Whatsapp-ChatBot De Pedidos Oriental Iniciado');
});

client.on('message', message => {
    console.log(message)
    if (message.id.remote == 'status@broadcast') {



    } else {
        if (message.type == 'chat') {
            if (message.body == 'BORRAR' || message.body == 'Borrar' || (message.body == 'borrar')) {
                sqlController.gestionContexto('', message.id.remote, 3)
                client.sendMessage(message.id.remote, 'TODO ELIMINADO, VACIE EL CHAT')
            }
            else {
                var mensajeInput = {
                    textMsg: message.body,
                    idChat: message.id.remote
                }
                var instancIaInput = message.id.id

                var paramsPetition = {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        intanciaWhatsapp: instancIaInput,
                        textMensajeReq: mensajeInput.textMsg,
                        idChat: mensajeInput.idChat,
                        idCanal: 1

                    })
                }



                const otrafuncion = async () => {
                    watsonResponse.EnviaWatson(paramsPetition, message.id.remote)
                }
                otrafuncion()
            }
        } else if (message.type == 'location') {
            var location = {
                from: message.from,
                location: message.location
            }
            var mensajeInput = {
                textMsg: 'UB-Estimado Chi Li Lee mi ubicacion es =' + location.location.latitude + '=' + location.location.longitude,
                idChat: message.id.remote
            }
            var instancIaInput = message.id.id

            var paramsPetition = {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    intanciaWhatsapp: instancIaInput,
                    textMensajeReq: mensajeInput.textMsg,
                    idChat: mensajeInput.idChat,
                    idCanal: 1

                })
            }



                const otrafuncion = async () => {
                    watsonResponse.EnviaWatson(paramsPetition, message.id.remote)
                }
                otrafuncion()
                console.log(location.location.latitude, location.location.longitude, 'Latitude y longitude');//aqui escribimos si necesitan usar funcion de localizacion. */ //aqui escribimos si necesitan usar funcion de localizacion.               

        }else {
            client.sendMessage(message.id.remote, 'Por el momento solo entiendo palabras escritas, proximamente podre entender otros tipos de mensajes')
          
        }

    }

});


async function enviamensaje(idchat, mensaje) {
    console.log(mensaje, 'llegada')
    let cadena = mensaje.text
    let img = 'image'
    let posicion = -1
    try {
        posicion = cadena.indexOf(img)
    } catch (error) {
        posicion = -1
    }


    if (posicion == 0) {
        console.log("noo")
        let cadena_nueva = cadena.split('^')
        mensaje.text = cadena_nueva[1]
        let url = cadena_nueva[0
        ]

        url = url.split(':')
        let urlNueva = url[1] + ':' + url[2]
        //console.log(urlNueva,'aaaa')
        let imgBase = ImgUrl(urlNueva, idchat, mensaje.title)

        //imgBase = imgBase.split(',')
        //console.log(imgBase[0])
        //const media = new MessageMedia('image/png',base64data)
        //client.sendMessage(idchat,media,{caption: mensaje.title})
    }


    if (mensaje.response_type == 'text') {

        if (mensaje.text != '' || mensaje.text != ' ') {
            await client.sendMessage(idchat, mensaje.text)
        }
        // await client.sendMessage(idchat, mensaje.text)
    }
    else {
        if (mensaje.response_type == 'option') {
            let respuesta = mensaje.title + '\n'
            mensaje.options.forEach(element => {
                respuesta = respuesta + element.label + '\n'
            });
            await client.sendMessage(idchat, respuesta)
        } else {
            if (mensaje.response_type == 'image') {

                await ImgUrl(mensaje.source, idchat, mensaje.title)
            }


        }

    }

}

client.initialize();


async function ImgUrl(urlPic, wsID) {
    console.log('2')
    var documentType
    var documentBase64
    var temporal1
    await fetch(urlPic)
        .then(res => {
            return res.blob()
        })
        .then((img64) => {
            documentType = img64.type
            temporal1 = img64.arrayBuffer()
        })

    documentBase64 = Buffer.from(await temporal1).toString('base64')
    const media = new MessageMedia(documentType, documentBase64, 'Catalogo Oriental')
    await client.sendMessage(wsID, media)

}

exports.enviamensaje = enviamensaje