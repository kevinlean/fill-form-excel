/* jshint esversion: 6 */
const http = require('http');

const axios = require('axios');
const Excel = require('exceljs');
const convert = require('xml-js');

const URL_DEV = 'http://172.16.1.33:7800/esb/service/SedeElectronica/';
const URL_PROD = 'http://172.16.1.127:7800/esb/service/SedeElectronica/';

const EXCEL_FILE_PATH = `${__dirname}/documents/MPN_MEC_Virtual.xlsx`;

const NUM_SOLICITUD_CELL = 'C';
const NUM_ORDEN_CELL = 'D';
const NUM_TRAMITE_CELL = 'E';
const ORGANIZACION_CELL = 'G';
const HASH_CELL = 'J'; /* Columna en la que se guardara el Hash generado */

let promises = [];
let rows = [];

let hasChanged = false;

// runExample();
const workbook = new Excel.Workbook();
workbook.xlsx.readFile(EXCEL_FILE_PATH)
    .then(() => {
        const worksheet = workbook.getWorksheet('Sheet1');
        const rowCount = worksheet.rowCount;

        const initial = 1;
        const limit = 10;

        console.log(`Current row range: ${initial} - ${limit}`);
        console.log(`Total number of rows: ${rowCount}`);

        for (let i = initial; i < limit; i++) {
            const row = worksheet.getRow(i);

            const numSolicitud = row.getCell(NUM_SOLICITUD_CELL).value;
            const numOrdenPago = row.getCell(NUM_ORDEN_CELL).value;
            const numTramite = row.getCell(NUM_TRAMITE_CELL).value;
            const organizacion = row.getCell(ORGANIZACION_CELL).value;
            const currentHash = row.getCell(HASH_CELL).value;

            /*
            ** Cuando es persona natural, el idTipoFormulario es 2.
            ** Si es establcimiento de comercio, el idTipoFormulario es 1
            */
            const idTipoFormulario = (organizacion.includes('2901') || organizacion.toLowerCase().includes('persona natural')) ? 2 : 1;
            const idTipoModelo = 2;
            const idTipoAplicativo = 1;

            /*
            ** Validar que exista un numero de orden de pago y que sea un numero mayor a 0,
            ** que exista un numero de tramite
            ** y que no se haya generado un hash.
            */
            if (parseInt(numOrdenPago) > 0 && numTramite && currentHash == null) {
                hasChanged = true;
                console.log(`numSolicitud: ${numSolicitud}`);
                console.log(`numOrdenPago: ${numOrdenPago}`);
                console.log(`numTramite: ${numTramite}`);
                console.log(`organizacion: ${organizacion}`);
                console.log(`idTipoFormulario: ${idTipoFormulario}`);
                console.log(`currentHash: ${currentHash}`);
                console.log('');

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
                                        <idTipoFormulario>${idTipoFormulario}</idTipoFormulario>
                                        <idTipoModelo>${idTipoModelo}</idTipoModelo>
                                        <idTipoAplicativo>${idTipoAplicativo}</idTipoAplicativo>
                                    </registrarFormulariosInDTO>
                                </data>
                            </ws:registrarFormularios>
                        </soapenv:Body>
                    </soapenv:Envelope>
                `;

                promises.push(sendRequest(URL_DEV, xml));
                rows.push(row);
            }
        }

        if (promises.length > 0) {
            const resultPromise = promises.reduce((promise, currentPromise, index) => {
                return delay(1000 * (index + 1)).then(_ =>
                    promise.then(_ =>
                        delay(1000).then(currentPromise.then(result => {
                            modifyCellValue(result.data, rows[index]);
                            const numSolicitud = rows[index].getCell(NUM_SOLICITUD_CELL).value;
                            console.log(`Executed promise with numSolicitud ${numSolicitud} at ${new Date()}.`);
                            console.log('');
                        }).catch(onError)
                    ))
                );
            }, Promise.resolve());

            resultPromise.then(_ => {
                console.log('All promises where executed.');
                if (hasChanged) {
                    workbook.xlsx.writeFile(EXCEL_FILE_PATH)
                        .then(() => {
                            console.log(`The file ${EXCEL_FILE_PATH} was modified successfully.`);
                        }).catch(onError);
                }
            });
        } else {
            console.log('There was not any row that matches all the conditions');
            console.log('Finished');
        }
    }).catch(onError);

function sendRequest(url, data) {
    const config = {
        url,
        data,
        method: 'POST',
        headers: {
            'Content-Type': 'text/xml'
        },
        responseType: 'text',
    };

    const promise = axios.request(config);
    return promise;
}

function modifyCellValue(parsedBody, row) {
    const numSolicitud = row.getCell(NUM_SOLICITUD_CELL).value;

    const jsonString = convert.xml2json(parsedBody, {
        compact: true
    });
    const json = JSON.parse(jsonString);

    const firstLevel = 'soapenv:Envelope';
    const secondLevel = 'soapenv:Body';
    const thirdLevel = 'NS1:registrarFormulariosResponse';

    let hasErrors = true;
    let hashText;

    try {
        hasErrors = json[firstLevel][secondLevel][thirdLevel].data.registrarFormulariosOutDTO.resultado.codigoError._text !== '0000';
        hashText = json[firstLevel][secondLevel][thirdLevel].metadata.transactionID._text;
    } catch (error) {
        console.error(`An error has ocurred getting the data: ${error}.`);
    }

    if (hasErrors === false && hashText) {
        console.log(`The hash for numSolicitud ${numSolicitud} is: ${hashText}.`);
        row.getCell(HASH_CELL).value = hashText;
    } else if (hasErrors) {
        console.error('Couldn\'t generate the hash, it was an error in the execution.');
    } else if (!hashText) {
        console.error(`Invalid Hash: ${hashText}.`);
    }
}

function delay(delay, value) {
    return new Promise(resolve => setTimeout(resolve.bind(null, value), delay));
}

function onError(error) {
    console.error(`Failed at ${new Date()}.`);
    console.error(error);
    console.log('');
}

function runExample() {
    const NUM_SOLICITUD = 578031;
    const NUM_ORDEN_PAGO = '0010382079';
    const NUM_TRAMITE = '000001700000611';
    const ID_TIPO_FORMULARIO = 1;
    const ID_TIPO_MODELO = 2;
    const ID_TIPO_APLICATIVO = 1;

    const TEST_XML = `
        <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ws="http://ws.recaudos.esb.ccb.org.co">
            <soapenv:Header/>
            <soapenv:Body>
                <ws:registrarFormularios>
                    <metadata/>
                    <data>
                        <registrarFormulariosInDTO>
                            <numSolicitud>${NUM_SOLICITUD}</numSolicitud>
                            <numOrdenPago>${NUM_ORDEN_PAGO}</numOrdenPago>
                            <numTramite>${NUM_TRAMITE}</numTramite>
                            <idTipoFormulario>${ID_TIPO_FORMULARIO}</idTipoFormulario>
                            <idTipoModelo>${ID_TIPO_MODELO}</idTipoModelo>
                            <idTipoAplicativo>${ID_TIPO_APLICATIVO}</idTipoAplicativo>
                        </registrarFormulariosInDTO>
                    </data>
                </ws:registrarFormularios>
            </soapenv:Body>
        </soapenv:Envelope>`;

    sendRequest(URL, TEST_XML)
        .then(result => {
            console.log(`Executed promise with numSolicitud ${NUM_SOLICITUD} at ${new Date()}.`);
            console.log('');
        }).catch(onError);
}