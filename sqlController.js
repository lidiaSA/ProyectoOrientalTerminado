
//API SOLO PARA GENERAR LA FUNCION DE BORRAR CONTEXTO


const mssql = require('mssql');
const sqlController = {};

sqlController.gestionContexto = async(contexto, idWhatsAppChat, opcion) =>{
    //opcion = 1 => registrar contexto
    //opcion = 2 => obtener contexto
    //opcion = 3 => eliminar contexto
    console.log(idWhatsAppChat)
    var config = {
        user: 'usergaia',
        password: 'Gaia2020',
        server: '192.168.100.3',
	port: 8282,
        database: 'ChatBotAyasaNissan' 
    };
    mssql.connect(config, function (err) {
    
        if (err) console.log(err);
        var query
        // create Request object
        var request = new mssql.Request();
        query = `exec [dbo].[sp_GestionarContexto] @_idChatWhatsApp = '${idWhatsAppChat}', @contexto = '${JSON.stringify(contexto)}', @opcion = '${opcion}'`
        // query to the database and get the records
        request.query(query, function (err, recordset) {
            
            if (err) console.log(err)

            // send records as a response
            
            
        });
    });

    
}

module.exports = sqlController
