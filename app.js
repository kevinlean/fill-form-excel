const http = require('http');

// const request = require('sync-request');
const rp = require('request-promise');
const Excel = require('exceljs');
const convert = require('xml-js');

const URL = 'http://172.16.1.33:7800/esb/service/SedeElectronica/';

const xlsxFile = `${__dirname}/documents/RECUPERA_MPN_VIRTUAL_2017-12-29.xlsx`;
const workbook = new Excel.Workbook();

const testXml = `
    <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ws="http://ws.recaudos.esb.ccb.org.co">
        <soapenv:Header/>
        <soapenv:Body>
            <ws:registrarFormularios>
                <metadata/>
                <data>
                    <registrarFormulariosInDTO>
                    <numSolicitud>578031</numSolicitud>
                    <numOrdenPago>0010382079</numOrdenPago>
                    <numTramite>000001700000611</numTramite>
                    <idTipoFormulario>1</idTipoFormulario>
                    <idTipoModelo>2</idTipoModelo>
                    <idTipoAplicativo>1</idTipoAplicativo>
                    </registrarFormulariosInDTO>
                </data>
            </ws:registrarFormularios>
        </soapenv:Body>
    </soapenv:Envelope>`;

// doRequest(URL, testXml);

const promises = [];
const rows = [];

workbook.xlsx.readFile(xlsxFile)
    .then(() => {
        const worksheet = workbook.getWorksheet('RECUPERA');
        const rowCount = worksheet.rowCount;

        console.log('Row Count', rowCount);

        for (let i = 2; i < 42; i++) {

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

                promises.push(doRequest(URL, xml));
                rows.push(row);
            }
        }

        promises.reduce((promiseChain, currentTask) => {
            return promiseChain.then(chainResults =>
                currentTask.then(currentResult => [...chainResults, currentResult])
            );
        }, Promise.resolve([])).then(arrayOfResults => {
            arrayOfResults.forEach((result, index) => modifyExcel(result, rows[index]));
            workbook.xlsx.writeFile(xlsxFile)
                .then(() => {
                    console.log('Done');
                }).catch(onError);
        }).catch(onError);
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
    console.log(new Date());

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

    const index = row.getCell(1).value;

    try {
        hasErrors = json[firstLevel][secondLevel][thirdLevel].data.registrarFormulariosOutDTO.resultado.codigoError._text !== '0000';
        hashText = json[firstLevel][secondLevel][thirdLevel].metadata.transactionID._text;
    } catch (error) {
        console.error('An error has ocurred getting the data:', error);
    }

    if (hasErrors === false && hashText) {
        console.log(`The hash for index ${index} is: ${hashText}`);
        row.getCell(10).value = hashText;
    } else if (hasErrors) {
        console.error('Couldn\'t generate the hash, it was an error in the execution');
    } else if (!hashText) {
        console.error('Invalid Hash:', hashText);
    }
}

function onError(error) {
    console.error(error);
}