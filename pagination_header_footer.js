import * as fs from "fs";
import { AlignmentType, Document, Footer, Header, Packer, PageBreak, PageNumber, Paragraph, TextRun } from "docx";

const doc = new Document({
    sections: [
        {
            properties: {
                titlePage: true,
            },
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [
                                new TextRun({
                                    children: [PageNumber.CURRENT],
                                }),
                            ],
                        }),
                    ],
                }),
                // precisa do first para adicionar paginação a primeira pagina
                first: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [
                                new TextRun({
                                    children: [PageNumber.CURRENT],
                                }),
                            ],
                        }),
                    ],
                }),
            },
            footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [
                                new TextRun({
                                    children: [PageNumber.CURRENT],
                                }),
                            ],
                        }),
                    ],
                }),
                // precisa do first para adicionar paginação a primeira pagina
                first: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.BOTH,
                            children: [
                                new TextRun({
                                    children: [PageNumber.CURRENT],
                                }),
                            ],
                        }),
                    ],
                }),
            },
            children: [
                new Paragraph({
                    children: [new TextRun("First Page"), new PageBreak()],
                }),
                new Paragraph("Second Page"),
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});