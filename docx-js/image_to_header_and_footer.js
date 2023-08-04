import * as fs from 'fs';
import { BorderStyle, Document, Footer, Header, ImageRun, Packer, Paragraph, Table, TableCell, TableRow, TableLayoutType, WidthType } from 'docx';

// Função para criar a tabela
const createTable = (imageBuffer) => {
    return new Table({
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Texto na Coluna 1")],
                    }),
                    new TableCell({
                        children: [
                            new Paragraph({
                                children: [
                                    new ImageRun({
                                        data: imageBuffer,
                                        transformation: {
                                            width: 50,
                                            height: 50,
                                        },
                                    }),
                                ],
                            }),
                        ],
                    }),
                    new TableCell({
                        children: [new Paragraph("Texto na Coluna 3")],
                    }),
                ],
            }),
        ],
        width: {
            size: 100,
            type: WidthType.PERCENTAGE,
        },
        layout: TableLayoutType.FIXED,
        borders: {
            top: {
                style: BorderStyle.NONE,
            },
            right: {
                style: BorderStyle.NONE,
            },
            bottom: {
                style: BorderStyle.NONE,
            },
            left: {
                style: BorderStyle.NONE,
            },
            insideHorizontal: {
                style: BorderStyle.NONE,
            },
            insideVertical: {
                style: BorderStyle.NONE,
            },
        },
    });
}

// Carregando a imagem
const imageBuffer = fs.readFileSync('logo.png');

const doc = new Document({
    sections: [
        {
            headers: {
                default: new Header({
                    children: [createTable(imageBuffer)],
                }),
            },
            footers: {
                default: new Footer({
                    children: [createTable(imageBuffer)],
                }),
            },
            children: [new Paragraph("Hello World, tudo bem?")],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("Meu Documento.docx", buffer);
});
