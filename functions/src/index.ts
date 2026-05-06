import * as functions from 'firebase-functions';
import * as admin from 'firebase-admin';
import { GoogleGenAI } from '@google/genai';

admin.initializeApp();

// Configurer l'API Key depuis les variables d'environnement de la fonction
const apiKey = process.env.GEMINI_API_KEY || functions.config().gemini?.key;
const ai = new GoogleGenAI({ apiKey: apiKey || '' });

export const generateProductEmbedding = functions.firestore
  .document('products/{productId}')
  .onWrite(async (change, context) => {
    // Si le produit est supprimé, on ne fait rien
    if (!change.after.exists) {
      return null;
    }

    const data = change.after.data();
    const previousData = change.before.data();

    const desc = data?.desc || '';
    const category = data?.category || '';

    const previousDesc = previousData?.desc || '';
    const previousCategory = previousData?.category || '';

    // Vérifier si desc ou category ont changé (ou si l'embedding est manquant)
    if (
      previousData &&
      desc === previousDesc &&
      category === previousCategory &&
      data?.embedding
    ) {
      return null; // Pas de changement nécessitant un nouvel embedding
    }

    const textToEmbed = `${category} ${desc}`.trim();
    
    if (!textToEmbed) {
      return null;
    }

    try {
      if (!apiKey) {
        console.error("Clé API Gemini manquante.");
        return null;
      }

      const response = await ai.models.embedContent({
        model: 'gemini-embedding-2-preview',
        contents: textToEmbed,
      });

      // Le SDK @google/genai formatte les embeddings dans response.embeddings
      const embedding = response.embeddings?.[0]?.values;

      if (embedding) {
        return change.after.ref.update({
          embedding: embedding,
        });
      }
    } catch (error) {
      console.error("Erreur lors de la génération de l'embedding:", error);
    }

    return null;
  });
