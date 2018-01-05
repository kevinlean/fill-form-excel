const http = require('http');

// const request = require('sync-request');
const rp = require('request-promise');
const Excel = require('exceljs');
const convert = require('xml-js');

const URLDEV = 'http://172.16.1.33:7800/esb/service/SedeElectronica/';
const URLPROD = 'http://172.16.1.127:7800/esb/service/SedeElectronica/';

const xlsxFile = `${__dirname}/documents/SinDireccion_2.xlsx`;
const workbook = new Excel.Workbook();

// const testXml = `
//     <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ws="http://ws.recaudos.esb.ccb.org.co">
//         <soapenv:Header/>
//         <soapenv:Body>
//             <ws:registrarFormularios>
//                 <metadata/>
//                 <data>
//                     <registrarFormulariosInDTO>
//                     <numSolicitud>578031</numSolicitud>
//                     <numOrdenPago>0010382079</numOrdenPago>
//                     <numTramite>000001700000611</numTramite>
//                     <idTipoFormulario>1</idTipoFormulario>
//                     <idTipoModelo>2</idTipoModelo>
//                     <idTipoAplicativo>1</idTipoAplicativo>
//                     </registrarFormulariosInDTO>
//                 </data>
//             </ws:registrarFormularios>
//         </soapenv:Body>
//     </soapenv:Envelope>`;

// doRequest(URL, testXml);

const promises = [];
const rows = [];

let hasChanged = false;

workbook.xlsx.readFile(xlsxFile)
    .then(() => {
        const worksheet = workbook.getWorksheet('Sheet1');
        const rowCount = worksheet.rowCount;

        // Total 692
        const initial = 1;
        const limit = 10;

        console.log('Current limit', limit);
        console.log('Row count', rowCount);

        for (let i = initial; i < limit; i++) {

            const row = worksheet.getRow(i);

            const index = row.getCell('A').value;
            const numSolicitud = row.getCell('B').value;
            const numOrdenPago = row.getCell('C').value;
            const numTramite = row.getCell('D').value;
            const currentHash = row.getCell('H').value;

            // Validar que el index sea mayor o igual a 0,
            // exista un numero de orden de pago,
            // un numero de tramite y que no se haya generado un hash
            if (index >= 0 && index !== null && numOrdenPago && numTramite && currentHash == null) {
                hasChanged = true;
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

                promises.push(doRequest(URLDEV, xml));
                rows.push(row);
            }
        }

        const resultPromise = promises.reduce((promise, currentPromise, index) => {
            return promise.then(_ =>
                currentPromise.then(result => modifyExcel(result, rows[index])).catch()
            );
        }, Promise.resolve());

        resultPromise.then(_ => {
            console.log('All promises where executed');
            if (hasChanged) {
                workbook.xlsx.writeFile(xlsxFile)
                    .then(() => {
                        console.log('The file was modified successfully');
                    }).catch(onError);
            }
        });

    }).catch(onError);

function doRequest(uri, body) {
    const options = {
        uri,
        body,
        method: 'POST',
        json: false,
        headers: {
            'Content-Type': 'text/xml'
        },
    };

    promise = rp(options);

    return promise;
}

function modifyExcel(parsedBody, row) {
    const index = row.getCell('A').value;
    console.log(`${index}: ${new Date()}`);

    const jsonString = convert.xml2json(parsedBody, {
        compact: true
    });
    const json = JSON.parse(jsonString);

    // console.log(json);

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
        row.getCell('H').value = hashText;
    } else if (hasErrors) {
        console.error('Couldn\'t generate the hash, it was an error in the execution');
    } else if (!hashText) {
        console.error('Invalid Hash:', hashText);
    }
}

function onError(error) {
    console.error(error);
}

// function sleep() {
//     return new Promise(resolve => setTimeout(() => console.log('llego'), 1000));
// }