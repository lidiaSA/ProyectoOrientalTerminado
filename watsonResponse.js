const watsonResponse = {};
var http = require("http");
var fetch = require('node-fetch');
var enviamensajeWs = require('./index.js')

watsonResponse.EnviaWatson = async(paramsPetition,idWhatsapp) =>{
    var uriApi = "http://127.0.0.1:8020/watson"
    var sendMessageRequest = {
      body: '',
      chatId: ''
    }
        var sendFileRequest

      fetch(uriApi,paramsPetition)
      .then(res => res.json())
      .then(resWatson => {
       (async () => {
            for (const element of resWatson) {
                let returned
                //console.log(element,'pertilente')
               await enviamensajeWs.enviamensaje(idWhatsapp,element)
                if(element.response_type == "text"){
                    //console.log(idWhatsapp,'parametros enviados')
                   
                    //console.log(element,'op1')
                } else if(element.response_type == "option"){
                    var texto = element.title + '\n'
                    element.options.forEach(element2 => {
                        texto = texto + '  • '+element2.label + '\n'
                    });

                    
                    // console.log(returned)

                    
                } else if(element.response_type == "image"){
                    //sendFileRequest.caption = element.title
                    
                    // console.log(returned)
                }
            }
        })()

        })
        .catch( err => {
            //si sudede erro, enviar a grupo de soporte el mensaje
            console.log(err)

            //res.send("error")
        })

    
}
module.exports = watsonResponse