import { GoogleGenAI, Type } from "@google/genai";
import { NextRequest, NextResponse } from "next/server";

const ai = new GoogleGenAI({
  apiKey: process.env.GEMINI_API_KEY,
  httpOptions: {
    headers: {
      'User-Agent': 'aistudio-build',
    }
  }
});

export async function POST(req: NextRequest) {
  try {
    const { image, mode } = await req.json();

    if (!image) {
      return NextResponse.json({ error: "No image provided" }, { status: 400 });
    }

    // Remove base64 prefix if present
    const base64Data = image.replace(/^data:image\/\w+;base64,/, "");

    const imagePart = {
      inlineData: {
        mimeType: "image/jpeg",
        data: base64Data,
      },
    };

    let prompt = "";
    let responseSchema = undefined;

    if (mode === "altea") {
      prompt = `Extract all text from this Altea screen or printed report. 
      Preserve the layout, line breaks, and spacing exactly as it appears. 
      This is airline terminal output (Cryptic or Content format). 
      Output ONLY the extracted text, no explanations.`;
    } else {
      // Table mode for BDO/SBH
      prompt = `Extract the baggage list from this document. 
      The document is an Excel sheet printout or a screen showing a table of baggage data.
      Convert the table into a structured JSON array of objects. 
      Identify columns such as Flight, Date, Tag (10-digit or airline format), Name, Weight, Status, and Remarks.
      Be precise with numbers and codes.
      If a column is missing, omit it. 
      Return only the JSON data.`;
      
      responseSchema = {
        type: Type.ARRAY,
        items: {
          type: Type.OBJECT,
          properties: {
            flightNo: { type: Type.STRING },
            receivedAt: { type: Type.STRING, description: "Date in ISO format if possible, or as text" },
            originalTag: { type: Type.STRING },
            name: { type: Type.STRING },
            weight: { type: Type.STRING },
            rushTag: { type: Type.STRING },
            remarks: { type: Type.STRING },
          }
        }
      };
    }

    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: [
        { inlineData: { mimeType: "image/jpeg", data: base64Data } },
        { text: prompt }
      ],
      config: responseSchema ? {
        responseMimeType: "application/json",
        responseSchema: responseSchema
      } : undefined
    });

    const text = response.text || "";
    
    if (responseSchema && text) {
        try {
            return NextResponse.json({ data: JSON.parse(text) });
        } catch (e) {
            return NextResponse.json({ text: text });
        }
    }

    return NextResponse.json({ text: text });
  } catch (error: any) {
    console.error("OCR Error:", error);
    return NextResponse.json({ error: error.message || "Failed to process image" }, { status: 500 });
  }
}
