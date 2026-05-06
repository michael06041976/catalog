import { collection, getDocs, doc, updateDoc } from "firebase/firestore";
import { db } from "./firebase";
import { GoogleGenAI } from "@google/genai";

export async function generateMissingEmbeddingsClientSide() {
  const apiKey = (import.meta as any).env.VITE_GEMINI_API_KEY || process.env.GEMINI_API_KEY;
  if (!apiKey) throw new Error("VITE_GEMINI_API_KEY non définie.");

  const ai = new GoogleGenAI({ apiKey });
  const productsSnap = await getDocs(collection(db, "products"));
  let updatedCount = 0;

  for (const docSnap of productsSnap.docs) {
    const data = docSnap.data();
    if (!data.embedding) {
      const textToEmbed = `${data.category || ''} ${data.desc || ''}`.trim();
      if (textToEmbed) {
        try {
          const response = await ai.models.embedContent({
            model: "gemini-embedding-2-preview",
            contents: textToEmbed,
          });
          const embedding = response.embeddings?.[0]?.values;
          if (embedding) {
            await updateDoc(doc(db, "products", docSnap.id), { embedding });
            updatedCount++;
          }
        } catch (err) {
          console.error("Erreur embedding produit:", docSnap.id, err);
        }
      }
    }
  }

  return updatedCount;
}
