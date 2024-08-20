const fs = require('fs');
const xlsx = require('node-xlsx');
const path = require('path');
const filePath = path.join(__dirname, 'file.xlsx');

const QUANTITY_OF_PERIODS = 2;
const QUANTITY_OF_FEMALE_PARTICIPANTS = 2;
const QUANTITY_OF_MALE_PARTICIPANTS = 2;

const ABBREVIATIONS = ['X1'];

const headerRow = [
    '',
    'NUMERO_ESTUDO',
    'ANO_ESTUDO',
    'NUMERO_DO_PARTICIPANTE',
    'NUMERO_INTERNACAO',
    'CLASSIFICACAO_MEDICAMENTO',
    'NUMERO_DA_CLASSIFICACAO',
    'CODIGO_DO_MEDICAMENTO',
    'SEXO_DO_PARTICIPANTE'
];

function populateParticipants(periods, femaleParticipants, maleParticipants) {
    const participants = [];

    for (let period = 0; period < periods; period++) {
        for (let female = 0; female < femaleParticipants; female++) {
            const participantOrder = female + 1;
            const participantPeriod = period + 1;
            const participantAbbreviation = ABBREVIATIONS[period] || ABBREVIATIONS[0];

            participants.push([
                '',
                (Math.floor(Math.random() * 9999)).toString(),
                '2024',
                participantOrder.toString(),
                participantPeriod.toString(),
                participantAbbreviation,
                '1',
                '0',
                'F'
            ])
        }

        for (let male = 0; male < maleParticipants; male++) {
            const participantOrder = femaleParticipants + male + 1;
            const participantPeriod = period + 1;
            const participantAbbreviation = ABBREVIATIONS[period] || ABBREVIATIONS[0];

            participants.push([
                '',
                (Math.floor(Math.random() * 9999)).toString(),
                '2024',
                participantOrder.toString(),
                participantPeriod.toString(),
                participantAbbreviation,
                '1',
                '0',
                'M'
            ])
        }
    }

    return participants;
}

const participants = populateParticipants(QUANTITY_OF_FEMALE_PARTICIPANTS, QUANTITY_OF_MALE_PARTICIPANTS, QUANTITY_OF_PERIODS);

var buffer = xlsx.build([{name: 'aleatorizacao', data: [
    [...headerRow],
    ...participants
]}]); // Returns a buffer

fs.writeFile(filePath, buffer, function(err) {
    if(err) {
        return console.log(err);
    }
    console.log("The file was saved!");
});
