import { initializeApp } from 'firebase/app';
import { getFirestore, doc, getDoc, setDoc, deleteDoc } from 'firebase/firestore';

const firebaseConfig = {
  apiKey: "AIzaSyDTdhgCf9YtDu1psnLrzuaKmGGseHZPc_g",
  authDomain: "robot-project-481517.firebaseapp.com",
  projectId: "robot-project-481517",
  storageBucket: "robot-project-481517.firebasestorage.app",
  messagingSenderId: "157099064510",
  appId: "1:157099064510:web:7b201ad5903910374f72ef"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app, "ai-studio-webhelpsystembui-5df430ed-4f09-47f3-847d-47d5224e5a06");

// Helper: Convert ArrayBuffer to Base64
export function arrayBufferToBase64(buffer: ArrayBuffer): string {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return window.btoa(binary);
}

// Helper: Convert Base64 to ArrayBuffer
export function base64ToArrayBuffer(base64: string): ArrayBuffer {
  const binaryString = window.atob(base64);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes.buffer;
}

export interface FirebaseDocData {
  docData: any; // parsed document structure
  fileBufferBase64: string | null; // docx file content
  theme: string;
  font: string;
  updatedAt: string;
}

// Save document globally to Firebase (using chunking to support >1MB files)
export async function saveDocumentToFirebase(
  docData: any,
  fileBuffer: ArrayBuffer | null,
  theme: string,
  font: string
): Promise<void> {
  try {
    const fileBufferBase64 = fileBuffer ? arrayBufferToBase64(fileBuffer) : null;
    
    // Create the full object to be saved
    const payload: FirebaseDocData = {
      docData,
      fileBufferBase64,
      theme,
      font,
      updatedAt: new Date().toISOString()
    };

    const serialized = JSON.stringify(payload);
    
    // Chunking settings (700,000 chars per document to remain well under 1MB Firestore limit)
    const chunkSize = 700000;
    const chunks: string[] = [];
    for (let i = 0; i < serialized.length; i += chunkSize) {
      chunks.push(serialized.substring(i, i + chunkSize));
    }
    
    const numChunks = chunks.length;
    const docRef = doc(db, 'global_document', 'current');
    
    // Retrieve previous numChunks so we can clean up any unused chunk documents
    let oldNumChunks = 0;
    try {
      const oldSnap = await getDoc(docRef);
      if (oldSnap.exists()) {
        const data = oldSnap.data();
        oldNumChunks = data.numChunks || 0;
      }
    } catch (e) {
      console.warn('Could not read old numChunks from Firestore:', e);
    }

    // 1. Write metadata
    await setDoc(docRef, {
      isChunked: true,
      numChunks,
      updatedAt: new Date().toISOString()
    });

    // 2. Write each chunk
    for (let i = 0; i < numChunks; i++) {
      const chunkRef = doc(db, 'global_document', 'current', 'chunks', `chunk_${i}`);
      await setDoc(chunkRef, {
        chunkIndex: i,
        data: chunks[i]
      });
    }

    // 3. Delete obsolete chunks if the new file is smaller than the old one
    const maxCleanIndex = Math.max(oldNumChunks, numChunks + 10);
    for (let i = numChunks; i < maxCleanIndex; i++) {
      const chunkRef = doc(db, 'global_document', 'current', 'chunks', `chunk_${i}`);
      try {
        await deleteDoc(chunkRef);
      } catch (e) {
        console.warn(`Failed to delete obsolete chunk_${i}:`, e);
      }
    }

    console.log(`Document successfully chunked and saved to Firebase across ${numChunks} chunks!`);
  } catch (error) {
    console.error('Error saving document to Firebase:', error);
    throw error;
  }
}

// Load global document from Firebase (reconstructing chunked parts)
export async function loadDocumentFromFirebase(): Promise<FirebaseDocData | null> {
  try {
    const docRef = doc(db, 'global_document', 'current');
    const docSnap = await getDoc(docRef);
    if (!docSnap.exists()) {
      return null;
    }
    
    const meta = docSnap.data();
    
    // Support non-chunked legacy format if present
    if (!meta.isChunked) {
      return meta as FirebaseDocData;
    }

    const numChunks = meta.numChunks || 0;
    if (numChunks === 0) {
      return null;
    }

    // Fetch all chunks in parallel
    const chunkPromises: Promise<any>[] = [];
    for (let i = 0; i < numChunks; i++) {
      const chunkRef = doc(db, 'global_document', 'current', 'chunks', `chunk_${i}`);
      chunkPromises.push(getDoc(chunkRef));
    }

    const chunkSnaps = await Promise.all(chunkPromises);
    let assembledString = '';
    
    for (let i = 0; i < numChunks; i++) {
      const chunkSnap = chunkSnaps[i];
      if (!chunkSnap.exists()) {
        throw new Error(`Missing document chunk: chunk_${i}`);
      }
      assembledString += chunkSnap.data().data;
    }

    const parsedData = JSON.parse(assembledString);
    return parsedData as FirebaseDocData;
  } catch (error) {
    console.error('Error loading document from Firebase:', error);
    return null;
  }
}
