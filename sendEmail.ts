import "dotenv/config";
import * as nodemailer from "nodemailer";
import * as pino from "pino";
import * as ExcelJS from "exceljs";

import { access } from "fs/promises";

const SENDER_EMAIL: string | undefined = process.env.SENDER_EMAIL;
const logger = pino({ level: process.env.PINO_LOG_LEVEL || "info" });
const SENDER_PASSWORD: string | undefined = process.env.SENDER_PASSWORD;
const RECEIVER_EMAIL: string | undefined = process.env.RECEIVER_EMAIL;
const SMTP_SERVER: string = "smtp.gmail.com";
const SMTP_PORT: number = 587;
const PROCESSED_DATA_FILE: string | undefined = process.env.PROCESSED_DATA_FILE;

async function sendLeadReportEmail() {
  logger.info("✨ Starting email sending process...");
  let recordsToSend = [];
  let emailBodyHtml = "";

  try {
    logger.info(`📚 Attempting to read processed data from '${PROCESSED_DATA_FILE}' for email body...`);
    const workbook = new ExcelJS.Workbook();
    if (!PROCESSED_DATA_FILE) {
      throw new Error(
        "PROCESSED_DATA_FILE is not defined in environment variables."
      );
    }
    await workbook.xlsx.readFile(PROCESSED_DATA_FILE);
    const processedDataSheet =
      workbook.getWorksheet("AI Processed Data") || workbook.getWorksheet(1);

    if (
      processedDataSheet &&
      processedDataSheet.rowCount &&
      processedDataSheet.rowCount > 1
    ) {
      const resultValues =
        (processedDataSheet.getRow(1)?.values) as string[] || [];
      const headers: string[] = resultValues.filter(Boolean);
      processedDataSheet.eachRow((row: ExcelJS.Row, rowNumber: number) => {
        if (rowNumber === 1) return;
        if (!row.values) return;

        let record = {};
        headers.forEach((header, colIndex) => {
          record[header] = row.getCell(colIndex + 1).value;
        });
        recordsToSend.push(record);
      });
      logger.info(`✅ Successfully read ${recordsToSend.length} records for email body.`);

      logger.info("🏗️ Constructing email body HTML...");
      
      emailBodyHtml = `
            <html>
            <head>
            <style>
              body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
              .container { max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px; background-color: #f9f9f9; }
              h2 { color: #0056b3; }
              .lead-item { background-color: #ffffff; border: 1px solid #eee; padding: 15px; margin-bottom: 20px; border-radius: 5px; }
              .lead-name { font-weight: bold; color: #333; font-size: 1.1em; margin-bottom: 5px; }
              .summary { margin-bottom: 10px; }
              .score { font-weight: bold; color: #008000; }
            </style>
            </head>
            <body>
              <div class="container">
                <h2>AI-Powered Lead Analysis for Digital Marketing</h2>
                <p>Hi Kydall,</p>
                <p>Here's the latest AI-generated lead analysis for potential digital marketing clients from New York:</p>
            `;

      recordsToSend.forEach((record) => {
        const businessName = record["Business Name"] || "N/A";
        const summary = record["AI Summary"] || "N/A";
        const leadQuality = record["Lead Quality Score"] || "N/A";

        emailBodyHtml += `
                <div class="lead-item">
                  <div class="lead-name">${businessName}</div>
                  <div class="summary"><strong>Summary:</strong> ${summary}</div>
                  <div class="score"><strong>Lead Quality:</strong> ${leadQuality}/10</div>
                </div>
                `;
      });

      emailBodyHtml += `
                <p>A full report is attached as an Excel file.</p>
                <p>Best regards,<br>Daniel\'s Automated Lead Generator</p>
              </div>
            </body>
            </html>
            `;
    } else {
      emailBodyHtml = `
            <html><body>
              <h2>AI-Powered Lead Analysis for Digital Marketing</h2>
              <p>Hi Kydall,</p>
              <p>No processed lead data available to display in the email body, but please find the report attached.</p>
              <p>Best regards,<br>Your Automated Lead Generator</p>
            </body></html>
            `;
      logger.warn("⚠️ No processed data found in XLSX for HTML email body, sending empty body.");
    }
  } catch (error) {
    console.error(
      `Error reading processed data from ${PROCESSED_DATA_FILE} for email body:`,
      error.message
    );
    emailBodyHtml = `
        <html><body>
          <h2>AI-Powered Lead Analysis Report</h2>
          <p>Hi Kydall,</p>
          <p>There was an error generating the detailed email body, but please find the processed report attached.</p>
          <p>Best regards,<br>Your Automated Lead Generator</p>
        </body></html>
        `;
  }

  logger.info(`🔍 Checking for processed data file at '${PROCESSED_DATA_FILE}' for attachment...`);
  let attachmentExists = false;
  try {
    await access(PROCESSED_DATA_FILE);
    attachmentExists = true;
  } catch (error) {
    logger.warn(`⚠️ Processed data file '${PROCESSED_DATA_FILE}' not found for attachment.`);
  }

  let transporter = nodemailer.createTransport({
    auth: {
      user: SENDER_EMAIL,
      pass: SENDER_PASSWORD,
    },
    host: SMTP_SERVER,
    port: SMTP_PORT,
    secure: false,
    tls: { rejectUnauthorized: false },
  });
  logger.info("📧 Nodemailer transporter created.");

  let mailOptions = {
    from: SENDER_EMAIL,
    to: RECEIVER_EMAIL,
    subject: `AI Lead Analysis Report - ${new Date().toLocaleDateString()}`,
    html: emailBodyHtml,
    attachments: attachmentExists
      ? [
          {
            filename: "Processed_Leads_Report.xlsx",
            path: PROCESSED_DATA_FILE,
            contentType:
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          },
        ]
      : [],
  };

  try {
    logger.info(`✉️ Attempting to send email to '${RECEIVER_EMAIL}'...`);
    let info = await transporter.sendMail(mailOptions);
    logger.info({ messageId: info.messageId }, "🎉 Email sent successfully!");
  } catch (error) {
    logger.error({ error }, "❌ Error sending email:");
    console.error(
      "Please ensure you're using an App Password for Gmail (if applicable) and correct SMTP settings."
    );
  }
}

sendLeadReportEmail();