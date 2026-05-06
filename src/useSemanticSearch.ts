import { useState, useCallback } from 'react';
import { GoogleGenAI } from '@google/genai';

const apiKey = (import.meta as any).env.VITE_GEMINI_API_KEY || process.env.GEMINI_API_KEY;
const ai = new GoogleGenAI({ apiKey });

export function useSemanticSearch() {
  const [isSearchingSemantic, setIsSearchingSemantic] = useState(false);

  const getEmbedding = useCallback(async (text: string): Promise<number[] | null> => {
    if (!text || !text.trim()) return null;
    
    setIsSearchingSemantic(true);
    try {
      const response = await ai.models.embedContent({
        model: 'gemini-embedding-2-preview',
        contents: text.trim(),
      });
      return response.embeddings?.[0]?.values || null;
    } catch (err) {
      console.error("Erreur lors de l'embedding de la requête:", err);
      return null;
    } finally {
      setIsSearchingSemantic(false);
    }
  }, []);

  return { getEmbedding, isSearchingSemantic };
}

export function cosineSimilarity(A: number[], B: number[]) {
  if (!A || !B || A.length !== B.length || A.length === 0) return 0;
  let dotProduct = 0;
  let normA = 0;
  let normB = 0;
  for (let i = 0; i < A.length; i++) {
    dotProduct += A[i] * B[i];
    normA += A[i] * A[i];
    normB += B[i] * B[i];
  }
  if (normA === 0 || normB === 0) return 0;
  return dotProduct / (Math.sqrt(normA) * Math.sqrt(normB));
}
