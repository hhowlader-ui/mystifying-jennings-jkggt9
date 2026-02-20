import { ExtractedData } from "../types";
import { GoogleGenAI } from "@google/genai";

// ---> PUT YOUR KEY RIGHT HERE <---
const apiKey = "AIzaSyBzrclQ7PMU5xh1C3HOUoyCNTlXfdEjT10";
const ai = new GoogleGenAI({ apiKey: apiKey });

// ... your export const parseBankStatementAuto code goes here ...

export const parseBankStatementAuto = async (
  file: File,
  contextHeaders?: string[]
): Promise<ExtractedData> => {
  console.log("Sending batch to Gemini OCR...", file.name);

  try {
    const arrayBuffer = await file.arrayBuffer();
    const base64String = btoa(
      new Uint8Array(arrayBuffer).reduce(
        (data, byte) => data + String.fromCharCode(byte),
        ""
      )
    );
    const mimeType = file.type || "application/pdf";

    const prompt = `
      Extract transaction data from this bank statement.
      Return ONLY valid JSON:
      {
        "headers": ["Date", "Description", "Money Out", "Money In", "Balance"],
        "rows": [["01 Jan 2026", "Example Payee", "50.00", "", "100.00"]]
      }
      Rules: Consistent dates, no currency symbols, empty strings for missing values. Raw JSON only.
    `;

    const response = await ai.models.generateContent({
      model: "gemini-2.5-flash-lite",
      contents: [prompt, { inlineData: { data: base64String, mimeType } }],
      config: {
        responseMimeType: "application/json",
        temperature: 0.0,
      },
    });

    const text = response.text;
    if (!text) throw new Error("No response text received from Gemini.");

    // --- NEW: Clean the text before parsing to prevent crash ---
    let cleanText = text.trim();
    if (cleanText.startsWith("```")) {
      cleanText = cleanText
        .replace(/^```(?:json)?\n?/i, "")
        .replace(/\n?```$/i, "");
    }

    const parsedData = JSON.parse(cleanText) as ExtractedData;

    parsedData.usage = {
      promptTokens: response.usageMetadata?.promptTokenCount || 0,
      responseTokens: response.usageMetadata?.candidatesTokenCount || 0,
    };

    return parsedData;
  } catch (error: any) {
    console.error("Full Gemini Error:", error);
    // --- NEW: Unmask the actual error from Google ---
    const errorMessage = error?.message || "Unknown Error occurred";
    throw new Error(`Gemini API Error: ${errorMessage}`);
  }
};
