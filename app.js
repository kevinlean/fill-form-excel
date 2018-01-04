const http = require('http');

const Excel = require('exceljs');
const convert = require('xml-js');

const host = '172.16.1.33';
const port = '7800';
const path = '/esb/service/SedeElectronica/';
const method = 'POST';

const postOptions = {
    host,
    port,
    path,
    method,
    headers: {
        'Content-Type': 'text/xml'
    }
};

const xlsxFile = `${__dirname}/documents/RECUPERA_MPN_VIRTUAL_2017-12-29.xlsx`;
const workbook = new Excel.Workbook();

let hashes = {};


// const testXml = `<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ws="http://ws.recaudos.esb.ccb.org.co">
//    <soapenv:Header/>
//    <soapenv:Body>
//       <ws:registrarFormularios>
//          <metadata/>
//          <data>
//             <registrarFormulariosInDTO>
//                <numSolicitud>578031</numSolicitud>
//                <numOrdenPago>0010382079</numOrdenPago>
//                <numTramite>000001700000611</numTramite>
//                <idTipoFormulario>1</idTipoFormulario>
//                <idTipoModelo>2</idTipoModelo>
//                <idTipoAplicativo>1</idTipoAplicativo>
//             </registrarFormulariosInDTO>
//          </data>
//       </ws:registrarFormularios>
//    </soapenv:Body>
// </soapenv:Envelope>`;

workbook.xlsx.readFile(xlsxFile)
    .then(function () {
        const worksheet = workbook.getWorksheet('RECUPERA');
        const rowCount = worksheet.rowCount;

        console.log('Row Count', rowCount);

        for (let i = 2; i < 30; i++) {

            const row = worksheet.getRow(i);

            const index = row.getCell(1).value;
            const numSolicitud = row.getCell(3).value;
            const numOrdenPago = row.getCell(5).value;
            const numTramite = row.getCell(7).value;
            const currentHash = row.getCell(10).value;

            // Validar que el index sea mayor a 0, 
            // exista un numero de solicitud, 
            // un numero de orden de pago, 
            // un numero de tramite y que no se haya generado un hash
            if (index > 0 && numSolicitud && numOrdenPago && numTramite && currentHash == null) {
                console.log('index:', index);
                console.log('numSolicitud:', numSolicitud);
                console.log('numOrdenPago:', numOrdenPago);
                console.log('numTramite:', numTramite);
                console.log('currentHash:', currentHash);
                
                const xml = `
                    <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ws="http://ws.recaudos.esb.ccb.org.co">
                        <soapenv:Header/>
                        <soapenv:Body>
                            <ws:registrarFormularios>
                                <metadata/>
                                <data>
                                    <registrarFormulariosInDTO>
                                    <numSolicitud>${numSolicitud}</numSolicitud>
                                    <numOrdenPago>${numOrdenPago}</numOrdenPago>
                                    <numTramite>${numTramite}</numTramite>
                                    <idTipoFormulario>1</idTipoFormulario>
                                    <idTipoModelo>2</idTipoModelo>
                                    <idTipoAplicativo>1</idTipoAplicativo>
                                    </registrarFormulariosInDTO>
                                </data>
                            </ws:registrarFormularios>
                        </soapenv:Body>
                    </soapenv:Envelope>
                `;


                const req = http.request(postOptions, function (res) {

                    // console.log('');
                    // console.log('Status: ' + res.statusCode);
                    // console.log('Headers: ' + JSON.stringify(res.headers));

                    res.setEncoding('utf8');
                    res.on('data', function (response) {
                        // console.log('Response: ' + response);

                        const jsonString = convert.xml2json(response, {
                            compact: true,
                            spaces: 4
                        });
                        const json = JSON.parse(jsonString);

                        // console.log('');
                        // console.log(jsonString);

                        const firstLevel = 'soapenv:Envelope';
                        const secondLevel = 'soapenv:Body';
                        const thirdLevel = 'NS1:registrarFormulariosResponse';

                        let hasErrors = true;
                        let hashText;

                        try {
                            hasErrors = json[firstLevel][secondLevel][thirdLevel].data.registrarFormulariosOutDTO.resultado.codigoError._text !== '0000';
                            hashText = json[firstLevel][secondLevel][thirdLevel].metadata.transactionID._text;
                        } catch (error) {
                            console.error('An error has ocurred getting the data:', error);
                        }

                        if (hasErrors === false && hashText) {
                            console.log(`The hash for index ${index} is: ${hashText}`);
                            hashes[index] = hashText;

                            // row.getCell(10).value = hashText;

                            // workbook.xlsx.writeFile(xlsxFile)
                            //     .then(function () {
                            //         console.log('Done');
                            //     });
                        } else if (hasErrors) {
                            console.log('Couldn\'t generate the hash, it was an error in the execution');
                        } else if (!hashText) {
                            console.log('Invalid Hash:', hashText);
                        }
                    });
                });

                // On error
                req.on('error', function (error) {
                    console.error('Problem with request: ', error);
                });

                // Post the data
                req.write(xml);
                req.end();
            }
        }
    });