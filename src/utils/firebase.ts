import { initializeApp } from 'firebase/app';
import { getFirestore, doc, getDoc, setDoc } from 'firebase/firestore';

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

// Save document globally to Firebase
export async function saveDocumentToFirebase(
  docData: any,
  fileBuffer: ArrayBuffer | null,
  theme: string,
  font: string
): Promise<void> {
  try {
    const fileBufferBase64 = fileBuffer ? arrayBufferToBase64(fileBuffer) : null;
    
    // To handle potential document size limits (1MB in Firestore),
    // we save docData and fileBuffer in separate docs if needed, or together.
    // Let's store them in a single document first. If docx is huge, it will alert.
    // Usually standard help documents are small (<500KB) and fit easily.
    const docRef = doc(db, 'global_document', 'current');
    
    await setDoc(docRef, {
      docData,
      fileBufferBase64,
      theme,
      font,
      updatedAt: new Date().toISOString()
    });
    console.log('Document successfully saved to Firebase!');
  } catch (error) {
    console.error('Error saving document to Firebase:', error);
    throw error;
  }
}

// Load global document from Firebase
export async function loadDocumentFromFirebase(): Promise<FirebaseDocData | null> {
  try {
    const docRef = doc(db, 'global_document', 'current');
    const docSnap = await getDoc(docRef);
    if (docSnap.exists()) {
      return docSnap.data() as FirebaseDocData;
    }
    return null;
  } catch (error) {
    console.error('Error loading document from Firebase:', error);
    return null;
  }
}
