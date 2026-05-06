import { collection, getDocs, doc, setDoc } from 'firebase/firestore';
import { ref, uploadString, getDownloadURL } from 'firebase/storage';
import { db, storage } from './firebase';

export async function migrateImagesFromBase64ToStorage() {
  console.log('Starting image migration...');
  try {
    const productsSnapshot = await getDocs(collection(db, 'products'));
    let migratedCount = 0;

    for (const d of productsSnapshot.docs) {
      const data = d.data();
      const img = data.img;

      if (img && typeof img === 'string' && img.startsWith('data:image')) {
        console.log(`Migrating image for product: ${d.id}`);
        // Extract the base64 string
        const base64Data = img.split(',')[1];
        // Extract mime type (e.g., data:image/jpeg;base64,... => image/jpeg)
        const mimeTypeMatch = img.match(/data:([a-zA-Z0-9]+\/[a-zA-Z0-9-.+]+).*,.*/);
        const mimeType = mimeTypeMatch ? mimeTypeMatch[1] : 'image/jpeg';
        
        const fileExt = mimeType.split('/')[1] || 'jpg';
        const imgRef = ref(storage, `products/${d.id}/${Date.now()}_migration.${fileExt}`);
        
        try {
          await uploadString(imgRef, base64Data, 'base64', { contentType: mimeType });
          const url = await getDownloadURL(imgRef);
          
          await setDoc(doc(db, 'products', d.id), { img: url }, { merge: true });
          console.log(`Successfully migrated product: ${d.id}`);
          migratedCount++;
        } catch (err) {
          console.error(`Failed to migrate product: ${d.id}`, err);
        }
      }
    }
    
    console.log(`Migration complete. Migrated ${migratedCount} products.`);
    return migratedCount;
  } catch (error) {
    console.error('Migration failed:', error);
    throw error;
  }
}
