import React, { useState, useRef, useMemo, useEffect } from 'react';
import { Upload, Download, Search, FileText, Menu, X, Github, ChevronRight, Palette, Type, Lock, Settings, LogOut, Info } from 'lucide-react';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { processDocx, ThemeColor } from './utils/docx';
import { DocData } from './types';
import { saveToDB, getFromDB } from './utils/db';
import { saveDocumentToFirebase, loadDocumentFromFirebase, base64ToArrayBuffer } from './utils/firebase';

export type FontOption = 'font-heebo' | 'font-rubik' | 'font-alef';

const FONTS: { id: FontOption, label: string }[] = [
  { id: 'font-heebo', label: 'Heebo' },
  { id: 'font-rubik', label: 'Rubik' },
  { id: 'font-alef', label: 'Alef' },
];

const THEME_OPTIONS: { id: ThemeColor, label: string, colors: string[] }[] = [
  { id: 'activePresenter', label: 'ActivePresenter', colors: ['bg-black', 'bg-white', 'bg-slate-600', 'bg-gray-200', 'bg-blue-500', 'bg-emerald-500', 'bg-orange-500', 'bg-amber-500', 'bg-purple-500', 'bg-rose-400'] },
  { id: 'classic', label: 'Classic', colors: ['bg-black', 'bg-white', 'bg-slate-600', 'bg-gray-200', 'bg-yellow-200', 'bg-orange-300', 'bg-red-400', 'bg-green-300', 'bg-cyan-300', 'bg-blue-400'] },
  { id: 'grayscale', label: 'Grayscale', colors: ['bg-black', 'bg-white', 'bg-black', 'bg-white', 'bg-gray-200', 'bg-gray-300', 'bg-gray-400', 'bg-gray-500', 'bg-gray-600', 'bg-gray-700'] },
  { id: 'blue', label: 'Blue', colors: ['bg-black', 'bg-white', 'bg-slate-800', 'bg-blue-100', 'bg-blue-600', 'bg-blue-500', 'bg-blue-400', 'bg-indigo-700', 'bg-slate-500', 'bg-stone-300'] },
  { id: 'blueGreen', label: 'Blue Green', colors: ['bg-black', 'bg-white', 'bg-slate-800', 'bg-blue-200', 'bg-blue-600', 'bg-blue-500', 'bg-sky-400', 'bg-emerald-600', 'bg-green-500', 'bg-lime-400'] },
  { id: 'green', label: 'Green', colors: ['bg-black', 'bg-white', 'bg-slate-900', 'bg-cyan-100', 'bg-green-600', 'bg-green-700', 'bg-emerald-700', 'bg-teal-600', 'bg-teal-700', 'bg-green-400'] },
  { id: 'greenYellow', label: 'Green Yellow', colors: ['bg-black', 'bg-white', 'bg-stone-800', 'bg-lime-200', 'bg-lime-500', 'bg-green-800', 'bg-green-600', 'bg-lime-400', 'bg-yellow-400', 'bg-yellow-300'] },
  { id: 'red', label: 'Red', colors: ['bg-black', 'bg-white', 'bg-stone-900', 'bg-orange-100', 'bg-red-600', 'bg-orange-500', 'bg-orange-400', 'bg-amber-700', 'bg-red-800', 'bg-stone-600'] },
  { id: 'redViolet', label: 'Red Violet', colors: ['bg-black', 'bg-white', 'bg-slate-900', 'bg-pink-200', 'bg-fuchsia-600', 'bg-purple-500', 'bg-rose-500', 'bg-fuchsia-400', 'bg-rose-400', 'bg-indigo-800'] },
  { id: 'yellow', label: 'Yellow', colors: ['bg-black', 'bg-white', 'bg-stone-600', 'bg-orange-100', 'bg-amber-400', 'bg-amber-500', 'bg-yellow-500', 'bg-amber-600', 'bg-stone-500', 'bg-orange-600'] },
  { id: 'yellowOrange', label: 'Yellow Orange', colors: ['bg-black', 'bg-white', 'bg-stone-700', 'bg-orange-200', 'bg-amber-500', 'bg-amber-600', 'bg-yellow-700', 'bg-orange-500', 'bg-orange-400', 'bg-orange-300'] },
];

export default function App() {
  const [docData, setDocData] = useState<DocData | null>(null);
  const [fileBuffer, setFileBuffer] = useState<ArrayBuffer | null>(null);

  const [theme, setTheme] = useState<ThemeColor>(() => {
    return (localStorage.getItem('theme') as ThemeColor) || 'activePresenter';
  });

  const [font, setFont] = useState<FontOption>(() => {
    return (localStorage.getItem('font') as FontOption) || 'font-heebo';
  });

  const [searchQuery, setSearchQuery] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [isSavingFirebase, setIsSavingFirebase] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [themeMenuOpen, setThemeMenuOpen] = useState(false);
  const [fontMenuOpen, setFontMenuOpen] = useState(false);
  const [isAdmin, setIsAdmin] = useState(false);
  const [showLogin, setShowLogin] = useState(false);
  const [password, setPassword] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (password === 'mos$gal%') { // Default simple password
      setIsAdmin(true);
      setShowLogin(false);
      setPassword('');
    } else {
      alert('סיסמה שגויה');
    }
  };

  const handleExitAdmin = async () => {
    if (docData) {
      setIsSavingFirebase(true);
      setError(null);
      try {
        await saveDocumentToFirebase(docData, fileBuffer, theme, font);
        console.log('Document successfully locked globally to Firebase');
      } catch (err: any) {
        console.error('Failed to lock document to Firebase:', err);
        setError('שגיאה בשמירת המסמך לענן: ' + (err.message || err));
      } finally {
        setIsSavingFirebase(false);
      }
    }
    setIsAdmin(false);
  };

  // Load document on initial mount
  useEffect(() => {
    const initializeDocument = async () => {
      setIsLoading(true);
      setError(null);
      
      try {
        // 1. Try loading from Firebase (Global source of truth)
        const globalDoc = await loadDocumentFromFirebase();
        if (globalDoc && globalDoc.docData) {
          console.log('Successfully loaded global document from Firebase');
          setDocData(globalDoc.docData);
          if (globalDoc.fileBufferBase64) {
            const buffer = base64ToArrayBuffer(globalDoc.fileBufferBase64);
            setFileBuffer(buffer);
            await saveToDB('fileBuffer', buffer);
          }
          if (globalDoc.theme) {
            setTheme(globalDoc.theme as ThemeColor);
            localStorage.setItem('theme', globalDoc.theme);
          }
          if (globalDoc.font) {
            setFont(globalDoc.font as FontOption);
            localStorage.setItem('font', globalDoc.font);
          }
          await saveToDB('docData', globalDoc.docData);
          setIsLoading(false);
          return;
        }
      } catch (err) {
        console.warn('Failed to load global document from Firebase, trying local fallbacks:', err);
      }
      
      // 2. Try to fetch default document.docx from public folder
      try {
        const response = await fetch('./document.docx');
        if (response.ok) {
          const arrayBuffer = await response.arrayBuffer();
          if (arrayBuffer.byteLength > 4) {
            const arr = new Uint8Array(arrayBuffer.slice(0, 4));
            const isZip = arr[0] === 0x50 && arr[1] === 0x4B && arr[2] === 0x03 && arr[3] === 0x04;
            if (isZip) {
              const data = await processDocx(arrayBuffer.slice(0), theme);
              setFileBuffer(arrayBuffer);
              setDocData(data);
              await saveToDB('fileBuffer', arrayBuffer);
              await saveToDB('docData', data);
              setIsLoading(false);
              return;
            } else {
              console.log('Fetched document.docx but it is not a valid zip/docx file (likely an HTML fallback page).');
            }
          }
        }
      } catch (e) {
        console.log('No default document.docx found, fallback to IndexedDB.', e);
      }

      // 3. Fallback to IndexedDB
      try {
        const savedBuffer = await getFromDB('fileBuffer');
        const savedDocData = await getFromDB('docData');
        if (savedBuffer && savedDocData) {
          setFileBuffer(savedBuffer);
          setDocData(savedDocData);
        }
      } catch (err) {
        console.error('Failed to load from IndexedDB:', err);
      } finally {
        setIsLoading(false);
      }
    };

    initializeDocument();
  }, []);


  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsLoading(true);
    setError(null);
    try {
      const arrayBuffer = await file.arrayBuffer();
      
      // Save original fileBuffer to IndexedDB
      try {
        await saveToDB('fileBuffer', arrayBuffer);
      } catch (err) {
        console.warn('Failed to save fileBuffer to IndexedDB:', err);
      }

      setFileBuffer(arrayBuffer);

      const data = await processDocx(arrayBuffer.slice(0), theme);
      setDocData(data);
      
      // Save docData to IndexedDB
      try {
        await saveToDB('docData', data);
      } catch (err) {
        console.warn('Failed to save docData to IndexedDB:', err);
      }

      setSearchQuery('');
    } catch (err: any) {
      setError(err.message || 'אירעה שגיאה בעיבוד הקובץ.');
    } finally {
      setIsLoading(false);
    }
  };

  const prevTheme = useRef<ThemeColor>(theme);

  useEffect(() => {
    // Save theme to localStorage
    localStorage.setItem('theme', theme);

    if (!fileBuffer) return;
    
    // Only reprocess if the theme actually changed from the previous value
    if (prevTheme.current !== theme) {
      prevTheme.current = theme;
      
      const reprocessDocument = async () => {
        setIsLoading(true);
        try {
          // We must pass a copy of the buffer because mammoth might consume/modify it
          const data = await processDocx(fileBuffer.slice(0), theme);
          setDocData(data);
          try {
            await saveToDB('docData', data);
          } catch (e) {
            console.warn('Failed to save docData to IndexedDB:', e);
          }
        } catch (err) {
          console.error('Failed to reprocess document with new theme:', err);
        } finally {
          setIsLoading(false);
        }
      };
      
      reprocessDocument();
    }
  }, [theme, fileBuffer]);

  useEffect(() => {
    localStorage.setItem('font', font);
  }, [font]);

  const handleDownloadMarkdown = async () => {
    if (!docData) return;
    
    try {
      const zip = new JSZip();
      
      // הוספת תיקיית docs בצורה מפורשת עבור תאימות של תוכנות פתיחת ZIP ומערכות הפעלה
      const docsFolder = zip.folder("docs");
      let navItems: string[] = [];

      const fullMd = docData.markdown;
      const lines = fullMd.split('\n');
      
      let currentFileTitle = 'ראשי';
      let currentFileName = 'index.md';
      let currentContent: string[] = [];
      let fileCounter = 1;
      
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (line.startsWith('## ')) {
          if (currentContent.length > 0 && currentContent.join('').trim().length > 0) {
            docsFolder?.file(currentFileName, currentContent.join('\n'));
            navItems.push(`  - "${currentFileTitle.replace(/"/g, '\\"')}": ${currentFileName}`);
          }
          
          currentFileTitle = line.replace(/^##\s*/, '').replace(/^#+\s*/, '').replace(/\s*#+$/, '').trim();
          const safeName = `section-${fileCounter}`;
          currentFileName = `${safeName}.md`;
          currentContent = [line];
          fileCounter++;
        } else {
          currentContent.push(line);
        }
      }
      
      if (currentContent.length > 0 && currentContent.join('').trim().length > 0) {
        docsFolder?.file(currentFileName, currentContent.join('\n'));
        navItems.push(`  - "${currentFileTitle.replace(/"/g, '\\"')}": ${currentFileName}`);
      }
      
      const hasIndex = navItems.some(nav => nav.endsWith(': index.md'));
      if (!hasIndex) {
        docsFolder?.file("index.md", "# ברוכים הבאים\n\nאנא בחרו נושא מהתפריט.");
        navItems.unshift(`  - "ראשי": index.md`);
      }

      const navString = navItems.length > 0 ? navItems.join('\n') : '  - "ראשי": index.md';

      // mkdocs.yml config
      let mkdocsColor = 'indigo';
      switch(theme) {
        case 'activePresenter': mkdocsColor = 'teal'; break;
        case 'classic': mkdocsColor = 'indigo'; break;
        case 'grayscale': mkdocsColor = 'grey'; break;
        case 'blue': mkdocsColor = 'blue'; break;
        case 'blueGreen': mkdocsColor = 'teal'; break;
        case 'green': mkdocsColor = 'green'; break;
        case 'greenYellow': mkdocsColor = 'light green'; break;
        case 'red': mkdocsColor = 'red'; break;
        case 'redViolet': mkdocsColor = 'purple'; break;
        case 'yellow': mkdocsColor = 'yellow'; break;
        case 'yellowOrange': mkdocsColor = 'orange'; break;
      }

      const mkdocsFont = font === 'font-rubik' ? 'Rubik' : font === 'font-alef' ? 'Alef' : 'Heebo';

      const mkdocsConfig = `site_name: My Documentation
site_url: https://your-username.github.io/your-repo-name/
nav:
${navString}
theme:
  name: material
  language: he
  font:
    text: ${mkdocsFont}
  features:
    - navigation.sections
    - toc.integrate
    - navigation.top
    - search.suggest
    - search.highlight
  palette:
    - media: "(prefers-color-scheme: light)"
      scheme: default
      primary: ${mkdocsColor}
      accent: ${mkdocsColor}
      toggle:
        icon: material/brightness-7 
        name: Switch to dark mode
    - media: "(prefers-color-scheme: dark)"
      scheme: slate
      primary: ${mkdocsColor}
      accent: ${mkdocsColor}
      toggle:
        icon: material/brightness-4
        name: Switch to light mode
`;
      zip.file("mkdocs.yml", mkdocsConfig);

      // GitHub actions workflow
      const workflowConfig = `name: ci
on:
  push:
    branches:
      - main
      - master
permissions:
  contents: write
jobs:
  deploy:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: 3.x
      - run: echo "cache_id=$(date --utc '+%V')" >> $GITHUB_ENV 
      - uses: actions/cache@v4
        with:
          key: mkdocs-material-\${{ env.cache_id }}
          path: .cache
          restore-keys: |
            mkdocs-material-
      - run: pip install mkdocs-material 
      - run: mkdocs gh-deploy --force
`;
      const githubFolder = zip.folder(".github");
      const workflowsFolder = githubFolder?.folder("workflows");
      workflowsFolder?.file("deploy.yml", workflowConfig);

      // קובץ .gitignore עבור MkDocs
      const gitignoreContent = `site/
.cache/
.DS_Store
*.log
`;
      zip.file(".gitignore", gitignoreContent);

      // קובץ README בתיקייה הראשית (אופציונלי) 
      const readmeContent = `# מסמך מיוצא

המסמך שלך יוצא בהצלחה.

## איך לראות את האתר ב-GitHub Pages?
1. העלה את כל הקבצים שחולצו מקובץ ה-ZIP למאגר (Repository) ב-GitHub.
2. ה-GitHub Actions ירוץ אוטומטית וייצור ענף בשם \`gh-pages\`.
3. גש ב-GitHub להגדרות המאגר: **Settings** -> **Pages**.
4. ודא שה-Source מוגדר ל-Deploy from a branch, והענף שנבחר הוא \`gh-pages\`.
5. הקישור לאתר יופיע בראש העמוד בהגדרות שם, בדרך כלל בפורמט:
   \`https://<username>.github.io/<repo-name>/\`

ראה את תוכן המסמך המקורי ב-[docs/index.md](docs/index.md).
`;
      zip.file("README.md", readmeContent);

      const blob = await zip.generateAsync({
        type: "blob",
        platform: "DOS",
        compression: "DEFLATE",
        compressionOptions: {
          level: 9
        }
      });
      saveAs(blob, "github-export.zip");
    } catch (error) {
      console.error("Error generating zip:", error);
      alert("אירעה שגיאה ביצירת קובץ ה-ZIP");
    }
  };

  const filteredSections = useMemo(() => {
    if (!docData) return [];
    if (!searchQuery.trim()) return docData.sections;
    
    const query = searchQuery.toLowerCase();
    return docData.sections.filter(sec => sec.title.toLowerCase().includes(query));
  }, [docData, searchQuery]);

  const sidebarTheme = useMemo(() => {
    switch (theme) {
      case 'activePresenter': return { bg: 'bg-slate-50', headerBg: 'bg-white', border: 'border-slate-200', text: 'text-slate-800', hoverBg: 'hover:bg-slate-100', hoverText: 'hover:text-slate-900', iconBg: 'bg-slate-600', focusRing: 'focus:ring-slate-500/20 focus:border-slate-500', groupFocus: 'group-focus-within:text-slate-600', dotBg: 'bg-slate-300', dotHover: 'group-hover:bg-slate-500', mainBg: 'bg-slate-50/50' };
      case 'classic': return { bg: 'bg-stone-50', headerBg: 'bg-white', border: 'border-stone-200', text: 'text-stone-800', hoverBg: 'hover:bg-stone-100', hoverText: 'hover:text-stone-900', iconBg: 'bg-stone-600', focusRing: 'focus:ring-stone-500/20 focus:border-stone-500', groupFocus: 'group-focus-within:text-stone-600', dotBg: 'bg-stone-300', dotHover: 'group-hover:bg-stone-500', mainBg: 'bg-stone-50/50' };
      case 'grayscale': return { bg: 'bg-gray-100', headerBg: 'bg-gray-50', border: 'border-gray-300', text: 'text-gray-900', hoverBg: 'hover:bg-gray-200', hoverText: 'hover:text-gray-900', iconBg: 'bg-gray-700', focusRing: 'focus:ring-gray-500/20 focus:border-gray-500', groupFocus: 'group-focus-within:text-gray-600', dotBg: 'bg-gray-300', dotHover: 'group-hover:bg-gray-600', mainBg: 'bg-gray-100/50' };
      case 'blue': return { bg: 'bg-blue-50/50', headerBg: 'bg-white', border: 'border-blue-100', text: 'text-blue-900', hoverBg: 'hover:bg-blue-100/80', hoverText: 'hover:text-blue-700', iconBg: 'bg-blue-600', focusRing: 'focus:ring-blue-500/20 focus:border-blue-500', groupFocus: 'group-focus-within:text-blue-500', dotBg: 'bg-blue-200', dotHover: 'group-hover:bg-blue-400', mainBg: 'bg-blue-50/30' };
      case 'blueGreen': return { bg: 'bg-sky-50/50', headerBg: 'bg-white', border: 'border-sky-100', text: 'text-sky-900', hoverBg: 'hover:bg-sky-100/80', hoverText: 'hover:text-sky-700', iconBg: 'bg-sky-600', focusRing: 'focus:ring-sky-500/20 focus:border-sky-500', groupFocus: 'group-focus-within:text-sky-500', dotBg: 'bg-sky-200', dotHover: 'group-hover:bg-sky-400', mainBg: 'bg-sky-50/30' };
      case 'green': return { bg: 'bg-emerald-50/50', headerBg: 'bg-white', border: 'border-emerald-100', text: 'text-emerald-900', hoverBg: 'hover:bg-emerald-100/80', hoverText: 'hover:text-emerald-700', iconBg: 'bg-emerald-600', focusRing: 'focus:ring-emerald-500/20 focus:border-emerald-500', groupFocus: 'group-focus-within:text-emerald-500', dotBg: 'bg-emerald-200', dotHover: 'group-hover:bg-emerald-400', mainBg: 'bg-emerald-50/30' };
      case 'greenYellow': return { bg: 'bg-lime-50/50', headerBg: 'bg-white', border: 'border-lime-100', text: 'text-lime-900', hoverBg: 'hover:bg-lime-100/80', hoverText: 'hover:text-lime-700', iconBg: 'bg-lime-600', focusRing: 'focus:ring-lime-500/20 focus:border-lime-500', groupFocus: 'group-focus-within:text-lime-500', dotBg: 'bg-lime-200', dotHover: 'group-hover:bg-lime-400', mainBg: 'bg-lime-50/30' };
      case 'red': return { bg: 'bg-red-50/50', headerBg: 'bg-white', border: 'border-red-100', text: 'text-red-900', hoverBg: 'hover:bg-red-100/80', hoverText: 'hover:text-red-700', iconBg: 'bg-red-600', focusRing: 'focus:ring-red-500/20 focus:border-red-500', groupFocus: 'group-focus-within:text-red-500', dotBg: 'bg-red-200', dotHover: 'group-hover:bg-red-400', mainBg: 'bg-red-50/30' };
      case 'redViolet': return { bg: 'bg-fuchsia-50/50', headerBg: 'bg-white', border: 'border-fuchsia-100', text: 'text-fuchsia-900', hoverBg: 'hover:bg-fuchsia-100/80', hoverText: 'hover:text-fuchsia-700', iconBg: 'bg-fuchsia-600', focusRing: 'focus:ring-fuchsia-500/20 focus:border-fuchsia-500', groupFocus: 'group-focus-within:text-fuchsia-500', dotBg: 'bg-fuchsia-200', dotHover: 'group-hover:bg-fuchsia-400', mainBg: 'bg-fuchsia-50/30' };
      case 'yellow': return { bg: 'bg-amber-50/50', headerBg: 'bg-white', border: 'border-amber-100', text: 'text-amber-900', hoverBg: 'hover:bg-amber-100/80', hoverText: 'hover:text-amber-700', iconBg: 'bg-amber-500', focusRing: 'focus:ring-amber-500/20 focus:border-amber-500', groupFocus: 'group-focus-within:text-amber-500', dotBg: 'bg-amber-200', dotHover: 'group-hover:bg-amber-400', mainBg: 'bg-amber-50/30' };
      case 'yellowOrange': return { bg: 'bg-orange-50/50', headerBg: 'bg-white', border: 'border-orange-100', text: 'text-orange-900', hoverBg: 'hover:bg-orange-100/80', hoverText: 'hover:text-orange-700', iconBg: 'bg-orange-600', focusRing: 'focus:ring-orange-500/20 focus:border-orange-500', groupFocus: 'group-focus-within:text-orange-500', dotBg: 'bg-orange-200', dotHover: 'group-hover:bg-orange-400', mainBg: 'bg-orange-50/30' };
      default: return { bg: 'bg-white', headerBg: 'bg-white', border: 'border-gray-200', text: 'text-gray-800', hoverBg: 'hover:bg-blue-50', hoverText: 'hover:text-blue-600', iconBg: 'bg-blue-600', focusRing: 'focus:ring-blue-500/20 focus:border-blue-500', groupFocus: 'group-focus-within:text-blue-600', dotBg: 'bg-gray-300', dotHover: 'group-hover:bg-blue-400', mainBg: 'bg-gray-50/50' };
    }
  }, [theme]);

  const scrollToSection = (id: string) => {
    const el = document.getElementById(id);
    if (el) {
      el.scrollIntoView({ behavior: 'smooth' });
    }
    setSidebarOpen(false);
  };

  return (
    <div className={`min-h-screen bg-gray-50 flex flex-col md:flex-row text-gray-900 ${font}`}>
      
      {/* Mobile Header */}
      <div className="md:hidden flex items-center justify-between bg-white border-b p-4 shadow-sm z-20">
        <div className="flex items-center gap-2 text-blue-600 font-bold text-lg">
          <FileText className="w-6 h-6" />
          <span>מערכת עזרה</span>
        </div>
        <button onClick={() => setSidebarOpen(true)} className="p-2 bg-gray-100 rounded-md">
          <Menu className="w-6 h-6 text-gray-700" />
        </button>
      </div>

      {/* Sidebar / Navigation */}
      <div className={`
        fixed inset-y-0 right-0 z-30 w-80 ${sidebarTheme.bg} border-l ${sidebarTheme.border} shadow-2xl transform transition-transform duration-300 ease-in-out
        md:relative md:translate-x-0 md:shadow-none flex flex-col
        ${sidebarOpen ? 'translate-x-0' : 'translate-x-full'}
      `}>
        <div className={`p-5 border-b ${sidebarTheme.border} flex items-center justify-between ${sidebarTheme.headerBg}`}>
          <div className={`flex items-center gap-3 ${sidebarTheme.text} font-bold text-xl tracking-tight`}>
            <div className={`${sidebarTheme.iconBg} p-2 rounded-lg text-white shadow-sm`}>
              <FileText className="w-5 h-5" />
            </div>
            <span>מערכת עזרה</span>
          </div>
          <button onClick={() => setSidebarOpen(false)} className="md:hidden p-2 text-gray-500 hover:bg-black/5 rounded-full transition-colors">
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Search */}
        <div className={`p-4 border-b ${sidebarTheme.border} ${sidebarTheme.mainBg}`}>
          <div className="relative group">
            <input
              type="text"
              placeholder="חיפוש בכותרות..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className={`w-full pl-4 pr-10 py-2.5 bg-white/80 border ${sidebarTheme.border} rounded-xl focus:outline-none focus:ring-2 ${sidebarTheme.focusRing} transition-all text-sm shadow-sm`}
              disabled={!docData}
            />
            <Search className={`absolute right-3 top-3 w-4 h-4 text-gray-400 ${sidebarTheme.groupFocus} transition-colors`} />
          </div>
        </div>

        {/* TOC List */}
        <div className="flex-1 overflow-y-auto p-4 custom-scrollbar">
          {!docData ? (
            <div className="flex flex-col items-center justify-center h-full text-center text-sm text-gray-400 space-y-3">
              <div className="w-16 h-16 bg-gray-50 border border-gray-100 rounded-full flex items-center justify-center mb-2 shadow-inner">
                <Menu className="w-8 h-8 text-gray-300" />
              </div>
              <p>ייבא קובץ DOCX<br/>כדי לראות את תוכן העניינים</p>
            </div>
          ) : filteredSections.length === 0 ? (
            <div className="text-center text-sm text-gray-500 mt-10">
              לא נמצאו תוצאות לחיפוש "{searchQuery}"
            </div>
          ) : (
            <ul className="space-y-1">
              {filteredSections.map((sec) => (
                <li key={sec.id} className="relative" style={{ paddingRight: `${(sec.level - 1) * 1}rem` }}>
                  {sec.level > 1 && (
                     <div className="absolute right-0 top-0 bottom-0 w-px bg-gray-200" style={{ right: `${(sec.level - 1) * 1 - 0.5}rem` }}></div>
                  )}
                  <button
                    onClick={() => scrollToSection(sec.id)}
                    className={`w-full text-right flex items-center text-sm py-2 px-3 rounded-lg transition-all duration-200
                      ${sec.level === 1 ? `font-semibold ${sidebarTheme.text} mt-4 mb-1 border-b ${sidebarTheme.border} pb-2` : 'text-gray-600'}
                      ${sidebarTheme.hoverText} ${sidebarTheme.hoverBg} hover:shadow-sm group
                    `}
                  >
                    {sec.level > 1 && (
                      <div className={`w-1.5 h-1.5 rounded-full ${sidebarTheme.dotBg} ml-2 ${sidebarTheme.dotHover} transition-colors shrink-0`}></div>
                    )}
                    <span className="truncate group-hover:translate-x-[-4px] transition-transform">{sec.title}</span>
                  </button>
                </li>
              ))}
            </ul>
          )}
        </div>
      </div>

      {/* Main Content */}
      <div className="flex-1 flex flex-col h-screen overflow-hidden">
        
        {/* Topbar Actions */}
        <header className="bg-white border-b px-6 py-4 flex flex-wrap items-center justify-between gap-4 shadow-sm z-10">
          <h1 className="text-xl font-semibold text-gray-800 hidden md:block">
            {docData ? 'צפייה במסמך' : 'ברוך הבא למערכת העזרה'}
          </h1>
          
          <div className="flex items-center gap-3 w-full md:w-auto">
            {/* Font Selector */}
            {docData && (
              <div className="relative">
                <button
                  onClick={() => setFontMenuOpen(!fontMenuOpen)}
                  className="flex items-center gap-2 px-3 py-2 bg-gray-100 text-gray-700 rounded-md hover:bg-gray-200 transition-colors text-sm font-medium border border-gray-200"
                  title="בחר פונט"
                >
                  <Type className="w-4 h-4" />
                  <span className="hidden md:inline">פונט</span>
                </button>
                
                {fontMenuOpen && (
                  <>
                    <div className="fixed inset-0 z-40" onClick={() => setFontMenuOpen(false)} />
                    <div className="absolute left-0 mt-2 w-48 bg-white border border-gray-200 rounded-md shadow-lg z-50 p-2 text-right">
                      <div className="text-xs text-gray-500 font-semibold mb-2 px-2">בחר פונט:</div>
                      <div className="flex flex-col gap-1">
                        {FONTS.map((f) => (
                          <button
                            key={f.id}
                            onClick={() => {
                              setFont(f.id);
                              setFontMenuOpen(false);
                            }}
                            className={`w-full text-right px-3 py-2 rounded-md text-sm transition-colors ${f.id} ${font === f.id ? 'bg-blue-50 text-blue-700 font-medium' : 'text-gray-700 hover:bg-gray-100'}`}
                          >
                            {f.label}
                          </button>
                        ))}
                      </div>
                    </div>
                  </>
                )}
              </div>
            )}

            {/* Theme Selector */}
            {docData && (
              <div className="relative">
                <button
                  onClick={() => setThemeMenuOpen(!themeMenuOpen)}
                  className="flex items-center gap-2 px-3 py-2 bg-gray-100 text-gray-700 rounded-md hover:bg-gray-200 transition-colors text-sm font-medium border border-gray-200"
                  title="בחר ערכת נושא"
                >
                  <Palette className="w-4 h-4" />
                  <span className="hidden md:inline">ערכת נושא</span>
                </button>
                
                {themeMenuOpen && (
                  <>
                    <div className="fixed inset-0 z-40" onClick={() => setThemeMenuOpen(false)} />
                    <div className="absolute left-0 mt-2 w-72 bg-white border border-gray-200 rounded-md shadow-lg z-50 p-2 text-right max-h-96 overflow-y-auto">
                      <div className="text-xs text-gray-500 font-semibold mb-2 px-2">בחר סגנון מובנה:</div>
                      <div className="flex flex-col gap-1">
                        {THEME_OPTIONS.map((t) => (
                          <button
                            key={t.id}
                            onClick={() => {
                              setTheme(t.id);
                              setThemeMenuOpen(false);
                            }}
                            className={`w-full flex items-center justify-between p-1.5 rounded-md border ${theme === t.id ? 'border-gray-400 bg-gray-50' : 'border-transparent hover:bg-gray-100'}`}
                          >
                            <div className="flex items-center gap-0.5 bg-gray-200 p-0.5 rounded border border-gray-300 shadow-inner">
                              {t.colors.map((colorClass, i) => (
                                <div key={i} className={`w-3.5 h-3.5 border border-black/10 ${colorClass}`}></div>
                              ))}
                            </div>
                            <span className="text-sm text-gray-700 ml-2 font-medium">{t.label}</span>
                          </button>
                        ))}
                      </div>
                    </div>
                  </>
                )}
              </div>
            )}

            <input
              type="file"
              accept=".docx"
              ref={fileInputRef}
              onChange={handleFileUpload}
              className="hidden"
            />
            {isAdmin && (
              <>
                <button
                  onClick={() => fileInputRef.current?.click()}
                  disabled={isLoading || isSavingFirebase}
                  className="flex-1 md:flex-none flex items-center justify-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors disabled:opacity-70 text-sm font-medium"
                >
                  <Upload className="w-4 h-4" />
                  {isLoading ? 'מייבא...' : 'ייבא קובץ Word'}
                </button>

                {docData && (
                  <button
                    onClick={async () => {
                      setIsSavingFirebase(true);
                      setError(null);
                      try {
                        await saveDocumentToFirebase(docData, fileBuffer, theme, font);
                        alert('המסמך נשמר וקובע במערכת בהצלחה עבור כל המשתמשים!');
                      } catch (err: any) {
                        alert('שגיאה בשמירה לענן: ' + (err.message || err));
                      } finally {
                        setIsSavingFirebase(false);
                      }
                    }}
                    disabled={isSavingFirebase || isLoading}
                    className="flex-1 md:flex-none flex items-center justify-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-md hover:bg-emerald-700 transition-colors disabled:opacity-70 text-sm font-medium"
                    title="שמור וקבע את המסמך בענן לכל המשתמשים"
                  >
                    <Settings className="w-4 h-4" />
                    {isSavingFirebase ? 'שומר בענן...' : 'שמור וקבע במערכת'}
                  </button>
                )}
              </>
            )}
            
            {!isAdmin ? (
              <button
                onClick={() => setShowLogin(true)}
                className="flex items-center justify-center p-2 text-gray-500 hover:text-gray-800 hover:bg-gray-100 rounded-md transition-colors"
                title="כניסה לניהול"
              >
                <Lock className="w-4 h-4" />
              </button>
            ) : (
              <button
                onClick={handleExitAdmin}
                disabled={isSavingFirebase}
                className="flex items-center justify-center gap-1.5 px-3 py-1 bg-green-100 text-green-800 hover:bg-green-200 text-xs font-bold rounded-full transition-colors disabled:opacity-70"
                title="יציאה מניהול ושמירה אוטומטית לענן"
              >
                <LogOut className="w-3 h-3" />
                {isSavingFirebase ? 'שומר ומקבע...' : 'מצב ניהול (יציאה ושמירה)'}
              </button>
            )}
          </div>
        </header>

        {/* Content Area */}
        <main className="flex-1 overflow-y-auto bg-gray-50 p-6 md:p-12">
          {error && (
            <div className="max-w-4xl mx-auto mb-6 bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded relative" role="alert">
              <span className="block sm:inline">{error}</span>
            </div>
          )}

          {isAdmin && (
            <div className="max-w-4xl mx-auto mb-6 bg-blue-50/80 border border-blue-200 text-blue-900 px-4 py-3 rounded-lg flex items-start gap-3 shadow-sm" dir="rtl">
              <Info className="w-5 h-5 text-blue-600 shrink-0 mt-0.5" />
              <div className="text-xs md:text-sm">
                <span className="font-bold">טיפ מנהל:</span> כדי לקבע את המסמך לצמיתות עבור כל מי שנכנס לאתר ב-GitHub Pages, שמור את הקובץ בשם <code className="bg-blue-100 text-blue-800 px-1 py-0.5 rounded font-mono">document.docx</code> בתוך תיקיית <code className="bg-blue-100 text-blue-800 px-1 py-0.5 rounded font-mono">public</code> במאגר ה-GitHub שלך.
              </div>
            </div>
          )}

          {!docData && !isLoading && !error && (
            <div className="max-w-2xl mx-auto mt-20 text-center bg-white p-10 rounded-xl shadow-sm border">
              <FileText className="w-16 h-16 text-blue-200 mx-auto mb-4" />
              <h2 className="text-2xl font-bold text-gray-800 mb-2">מערכת עזרה ויצירת תיעוד</h2>
              <p className="text-gray-600 mb-6 leading-relaxed">
                מערכת זו מאפשרת לך לייבא קבצי Word (בפורמט DOCX), להציג אותם בצורה נוחה עם סרגל ניווט וחיפוש,
                ולהמיר אותם לקובץ Markdown (MD) שמתאים לשימוש כתיעוד ב-GitHub.
              </p>
              {isAdmin ? (
                <button
                  onClick={() => fileInputRef.current?.click()}
                  className="inline-flex items-center gap-2 px-6 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors shadow-sm font-medium"
                >
                  <Upload className="w-5 h-5" />
                  בחר קובץ להתחלה
                </button>
              ) : (
                <button
                  onClick={() => setShowLogin(true)}
                  className="inline-flex items-center gap-2 px-6 py-3 bg-gray-100 text-gray-700 rounded-lg hover:bg-gray-200 transition-colors shadow-sm font-medium"
                >
                  <Lock className="w-5 h-5" />
                  כניסה למצב ניהול
                </button>
              )}
            </div>
          )}

          {isLoading && (
            <div className="flex flex-col items-center justify-center h-64">
              <div className="w-12 h-12 border-4 border-blue-200 border-t-blue-600 rounded-full animate-spin"></div>
              <p className="mt-4 text-gray-500 font-medium">מעבד את המסמך, אנא המתן...</p>
            </div>
          )}

          {docData && (
            <div 
              className="max-w-4xl mx-auto bg-white p-8 md:p-12 rounded-xl shadow-sm border doc-content"
              dangerouslySetInnerHTML={{ __html: docData.html }}
            />
          )}
        </main>
      </div>

      {/* Overlay for mobile sidebar */}
      {sidebarOpen && (
        <div 
          className="fixed inset-0 bg-black bg-opacity-50 z-20 md:hidden"
          onClick={() => setSidebarOpen(false)}
        />
      )}

      {/* Login Modal */}
      {showLogin && (
        <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl p-6 w-full max-w-sm" dir="rtl">
            <div className="flex items-center justify-between mb-6">
              <h3 className="text-xl font-bold text-gray-800 flex items-center gap-2">
                <Lock className="w-5 h-5 text-blue-600" />
                כניסה לניהול
              </h3>
              <button onClick={() => setShowLogin(false)} className="text-gray-400 hover:text-gray-600 transition-colors">
                <X className="w-5 h-5" />
              </button>
            </div>
            
            <form onSubmit={handleLogin} className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">סיסמה</label>
                <input
                  type="password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  className="w-full px-4 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  placeholder="הזן סיסמה..."
                  autoFocus
                />
                <p className="text-xs text-gray-500 mt-2">הזן סיסמת מנהל</p>
              </div>
              
              <button
                type="submit"
                className="w-full py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors font-medium"
              >
                היכנס
              </button>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
