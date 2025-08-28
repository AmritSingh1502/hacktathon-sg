import express from 'express';
import { Document, Packer, Paragraph } from 'docx';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';

const app = express();
app.use(express.json());

interface EmailData {
  sender?: string;
  subject?: string;
  date?: string;
  actionItems?: string;
}

app.post('/process-email', async (req, res) => {
  try {
    const emailData: EmailData = req.body;

    const structured = {
      sender: emailData.sender ?? 'unknown@example.com',
      subject: emailData.subject ?? 'No Subject',
      date: emailData.date ?? new Date().toISOString().slice(0, 10),
      actionItems: emailData.actionItems ?? 'No action items',
    };

    // Create Excel file
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Email Data');
    worksheet.columns = [
      { header: 'Sender', key: 'sender', width: 30 },
      { header: 'Subject', key: 'subject', width: 30 },
      { header: 'Date', key: 'date', width: 15 },
      { header: 'Action Items', key: 'actionItems', width: 40 },
    ];
    worksheet.addRow(structured);

    const outputDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir);
    }

    const excelPath = path.join(outputDir, 'email_data.xlsx');
    await workbook.xlsx.writeFile(excelPath);

    // Create Word document
    const doc = new Document({
      sections: [
        {
          children: [
            new Paragraph(`Sender: ${structured.sender}`),
            new Paragraph(`Subject: ${structured.subject}`),
            new Paragraph(`Date: ${structured.date}`),
            new Paragraph(`Action Items: ${structured.actionItems}`),
          ],
        },
      ],
    });

    const wordBuffer = await Packer.toBuffer(doc);
    const wordPath = path.join(outputDir, 'email_data.docx');
    fs.writeFileSync(wordPath, wordBuffer);

    // Create text file
    const textPath = path.join(outputDir, 'email_data.txt');
    fs.writeFileSync(textPath, JSON.stringify(structured, null, 2));

    res.json({
      message: 'Files generated successfully',
      files: {
        excel: 'output/email_data.xlsx',
        word: 'output/email_data.docx',
        text: 'output/email_data.txt',
      },
    });
  } catch (error) {
    console.error('Error processing email data:', error);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
