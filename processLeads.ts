import "dotenv/config";
import * as ExcelJS from "exceljs";
import { GoogleGenerativeAI } from "@google/generative-ai";
import * as pino from 'pino';

const RAW_DATA_FILE: string | undefined = process.env.RAW_DATA_FILE;
const logger = pino({ transport: { target: 'pino-pretty' } });
const PROCESSED_DATA_FILE: string | undefined = process.env.PROCESSED_DATA_FILE;
const GEMINI_API_KEY: string | undefined = process.env.GEMINI_API_KEY;

const genAI = new GoogleGenerativeAI(GEMINI_API_KEY!);
const model = genAI.getGenerativeModel({ model: "gemini-2.5-flash-preview-05-20" });

async function processLeads() {
  logger.info("🚀 Starting Lead Processing...");
  const workbook = new ExcelJS.Workbook();
  let rawRecords: { [key: string]: any }[] = [];

  logger.info(
    `📚 Reading raw data from '${RAW_DATA_FILE}'...`
  );

  if (!RAW_DATA_FILE) {
    logger.error("RAW_DATA_FILE environment variable is not set.");
    return;
  }
  try {
    await workbook.xlsx.readFile(RAW_DATA_FILE!);
    const rawDataSheet =
      workbook.getWorksheet("Raw Data") || workbook.getWorksheet(1);

    if (!rawDataSheet) {
      throw new Error(`Worksheet 'Raw Data' not found in ${RAW_DATA_FILE!}.`);
    }

    logger.info("✅ Raw data sheet found. Reading headers...");
    const resultValues = rawDataSheet.getRow(1).values as (
      | string
      | number
      | boolean
      | Date
      | undefined
    )[];
    const headers: (string | number | boolean | Date | undefined)[] =
      resultValues.filter(Boolean);
    if (headers.length === 0) {
      logger.error("❌ No headers found in the raw data sheet.");
      return;
    }
    logger.info(
      `📋 Headers read: ${headers.map((h) => `'${h}'`).join(", ")}`
    );
    rawDataSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      let record = {};
      headers.forEach((header, colIndex) => {
        record[header as string] = row.getCell(colIndex + 1).value;
      });
      rawRecords.push(record);
    });

    logger.info(`✅ Finished reading ${rawRecords.length} records from raw data.`);
  } catch (error: any) {
    logger.error(
      { file: RAW_DATA_FILE, error: error.message },
      `❌ Error reading raw data from ${RAW_DATA_FILE}`
    );
    return;
  }

  const processedResults = [];

  logger.info(`🧠 Starting AI processing for ${rawRecords.length} records...`);

  let processedCount = 0;
  for (const record of rawRecords) {
    processedCount++;
    const businessName: string = record["Company"];
    const keyInfo: string = record["About"];

    if (!keyInfo) {
      logger.info(`Skipping '${businessName}' due to missing 'About'.`);
      processedResults.push([
        businessName,
        "Skipped: Missing AI information",
        "0",
      ]);
      continue;
    }

    const prompt = `
        You are an expert digital marketing analyst for a top agency. Your task is to analyze a real estate agency and determine its potential as a lead for digital marketing services.

        **Business Name:** ${businessName}
        **Key Information about the Business (from their website/public sources):**
        ${keyInfo}

        Based on the information above, provide the following in a structured JSON format. Ensure the JSON is valid and only contains the specified keys.

        {
          "summary": "Summarize what this real estate agency does in exactly two sentences. Then, explain concisely why they would be a good lead for a digital marketing agency, specifically highlighting potential needs they might have (e.g., need for better SEO, social media presence, lead generation campaigns, website redesign, content marketing). This second part of the sentence should also be concise.",
          "lead_quality_score": "Rate the lead quality from 1 to 10 for a digital marketing agency, where 1 is very low potential and 10 is very high potential. Consider factors like their current online presence (implied from the info), market competitiveness, and clarity of services."
        }
        `;
    logger.info(`✨ Processing record ${processedCount}/${rawRecords.length}: ${businessName}...`);
    logger.debug({ prompt }, "AI Prompt:");
    try {
      const result = await model.generateContent(prompt);
      const aiOutputText: string = result.response.text();
      logger.debug({ aiOutputText }, "Raw AI Output:");
      const cleanedAiOutputText = aiOutputText
        .replace("```json", "")
        .replace("```", "")
        .trim();
      const aiData = JSON.parse(cleanedAiOutputText);

      const summary = aiData.summary || "N/A";
      const leadQualityScore = aiData.lead_quality_score || "N/A";

      processedResults.push([businessName, summary, leadQualityScore]);
      logger.info(`--- Finished processing ${businessName} - Score: ${leadQualityScore} ---`);
    } catch (error) {
      logger.error({ businessName, error: error.message }, `--- Error processing ${businessName} ---`);
      processedResults.push([businessName, `Error: ${error.message}`, "0"]);
    }
  }

  const processedWorkbook = new ExcelJS.Workbook();
  console.log(`\n--- Writing processed data to '${PROCESSED_DATA_FILE}' ---`);
  const processedSheet = processedWorkbook.addWorksheet("AI Processed Data");

  processedSheet.addRow(["Business Name", "AI Summary", "Lead Quality Score"]);

  processedResults.forEach((row) => {
    processedSheet.addRow(row);
  });

  try {
    await processedWorkbook.xlsx.writeFile(PROCESSED_DATA_FILE);
    logger.info(`--- Successfully wrote processed data to '${PROCESSED_DATA_FILE}' ---`);
  } catch (error) {
    logger.error(
      { file: PROCESSED_DATA_FILE, error: error.message },
      error.message
    );
  }

  logger.info(`
-------------------------------------------------------------------
             Lead Processing Complete!
-------------------------------------------------------------------
  `);
}

processLeads();
