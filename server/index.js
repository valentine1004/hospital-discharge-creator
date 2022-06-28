const docx = require("docx")
const express = require("express");
const cors = require('cors');
var bodyParser = require('body-parser')

const app = express();
const port = 8000;

app.use(cors());
app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());

const { Document, Packer, Paragraph, TextRun } = docx;

app.post("/generate-doc", async (req, res) => {

    const { 
        cardId, 
        name, 
        birthDay, 
        address,
        jobTitle,
        startTreatment,
        endTreatment,
        diagnosis
    } = req.body;

    const doc = new Document({
        sections: [{
            properties: {},
            children: [
                new Paragraph({
                    children: [
                        new TextRun({ 
                            text: "Виписка", 
                            bold: true,
                            allCaps: true,
                            size: 24
                        }),
                    ],
                    alignment: 'center',
                    spacing: {
                        after: 100,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({ 
                            text: `із медичної картки стаціонарного хворого № ${cardId}`, 
                            bold: true,
                            size: 24
                        }),
                    ],
                    alignment: 'center',
                    spacing: {
                        after: 300,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({ 
                            text: `1. Прізвище, ім’я по батькові хворого: ${name}`, 
                            bold: true,
                            size: 24
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({ 
                            text: '2. Дата народження: ', 
                            bold: true,
                            size: 24
                        }),
                        new TextRun({
                            text: `${birthDay}р.`,
                            size: 24
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({ 
                            text: '3. Домашня адреса: ', 
                            bold: true,
                            size: 24
                        }),
                        new TextRun({
                            text: `${address}`,
                            size: 24
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({ 
                            text: '4. Місце роботи, посада: ', 
                            bold: true,
                            size: 24
                        }),
                        new TextRun({
                            text: `${jobTitle}`,
                            size: 24
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({ 
                            text: '5. Знаходився(лась) на лікуванні з ', 
                            bold: true,
                            size: 24
                        }),
                        new TextRun({
                            text: `${startTreatment}р.`,
                            size: 24,
                            underline: {
                                type: 'single',
                            },
                        }),
                        new TextRun({
                            text: ' по ',
                            bold: true,
                            size: 24
                        }),
                        new TextRun({
                            text: `${endTreatment ? endTreatment : '00.00.2022'}р.`,
                            size: 24,
                            underline: {
                                type: 'single',
                            },
                            highlight: endTreatment ? "white" : "yellow",
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({ 
                            text: '6. Повний діагноз (основне захворювання, супутні захворювання та ускладнення): ', 
                            bold: true,
                            size: 24
                        }),
                        new TextRun({
                            text: `${diagnosis}`,
                            size: 24
                        }),
                    ],
                }),
            ],
        }],
    });

    const b64string = await Packer.toBase64String(doc);
    
    res.setHeader('Content-Disposition', 'attachment; filename=My Document.docx');
    res.send(Buffer.from(b64string, 'base64'));
})

app.listen(port, () => {
    console.log(`Server listening on port ${port}`)
});