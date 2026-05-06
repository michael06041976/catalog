import { useState, useCallback } from 'react';
import { GoogleGenAI, Type } from '@google/genai';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

export interface GeminiProductResult {
  sku: string;
  desc: string;
  category: string;
  supplier: string;
}

export function useGeminiProduct(imageBase64: string | null) {
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const analyzeImage = useCallback(async (): Promise<GeminiProductResult | null> => {
    if (!imageBase64) {
      setError('אין תמונה לניתוח');
      return null;
    }
    
    setIsAnalyzing(true);
    setError(null);
    try {
      const base64Data = imageBase64.split(',')[1] || imageBase64;
      const mimeType = imageBase64.match(/data:(.*?);base64/)?.[1] || 'image/jpeg';

      const response = await ai.models.generateContent({
        model: 'gemini-2.0-flash',
        contents: [
          {
            role: 'user',
            parts: [
              { text: 'Analyze this product image and extract the following details: SKU/Barcode, Description, Category, and Supplier/Brand. The application uses Hebrew/French, so please provide the description and category in Hebrew (preferred) or French. Identify the manufacturer or brand for supplier. If you cannot determine a field, return an empty string. The response MUST be a valid JSON object.' },
              { inlineData: { data: base64Data, mimeType } }
            ]
          }
        ],
        config: {
            responseMimeType: 'application/json',
            responseSchema: {
                type: Type.OBJECT,
                properties: {
                    sku: { type: Type.STRING, description: 'Product identifier, barcode, or SKU visibly read from the image' },
                    desc: { type: Type.STRING, description: 'A short description or name of the product in Hebrew or French' },
                    category: { type: Type.STRING, description: 'Product category in Hebrew or French' },
                    supplier: { type: Type.STRING, description: 'Supplier, brand name or manufacturer' },
                },
            }
        }
      });

      if (!response.text) {
          throw new Error('No response from Gemini');
      }

      return JSON.parse(response.text) as GeminiProductResult;
    } catch (err: any) {
      setError('שגיאה בניתוח התמונה - Gemini לא הצליח לזהות את המוצר');
      console.error('Gemini API Error:', err);
      return null;
    } finally {
      setIsAnalyzing(false);
    }
  }, [imageBase64]);

  return { analyzeImage, isAnalyzing, error };
}
