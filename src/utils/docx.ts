import mammoth from 'mammoth';
import TurndownService from 'turndown';
import { gfm } from 'turndown-plugin-gfm';
import { DocData, DocSection } from '../types';

function findFirstTextNode(node: Node): Node | null {
  if (node.nodeType === 3) return node; // 3 is Node.TEXT_NODE
  for (let i = 0; i < node.childNodes.length; i++) {
    const found = findFirstTextNode(node.childNodes[i]);
    if (found) return found;
  }
  return null;
}

function findLastTextNode(node: Node): Node | null {
  if (node.nodeType === 3) return node; // 3 is Node.TEXT_NODE
  for (let i = node.childNodes.length - 1; i >= 0; i--) {
    const found = findLastTextNode(node.childNodes[i]);
    if (found) return found;
  }
  return null;
}

export type ThemeColor = 'activePresenter' | 'classic' | 'grayscale' | 'blue' | 'blueGreen' | 'green' | 'greenYellow' | 'red' | 'redViolet' | 'yellow' | 'yellowOrange';

export async function processDocx(arrayBuffer: ArrayBuffer, theme: ThemeColor = 'activePresenter'): Promise<DocData> {
  let html = '';
  try {
    const result = await mammoth.convertToHtml({ arrayBuffer });
    html = result.value;
  } catch (mammothError: any) {
    console.error('Mammoth failed to parse file:', mammothError);
    const msg = mammothError?.message || '';
    if (msg.includes('zip') || msg.includes('word/document.xml') || msg.includes('Can\'t find') || msg.includes('signature')) {
      throw new Error('קובץ לא תקין או שאינו בפורמט DOCX. שים לב: המערכת תומכת בקבצי Word מסוג DOCX בלבד (לא קבצי .doc ישנים, PDF, תמונות או קבצים פגומים). אנא שמור את הקובץ מחדש כ-Word Document (.docx) ונסה שוב.');
    }
    throw new Error(`שגיאה בקריאת קובץ ה-Word: ${msg || 'ודא שהקובץ תקין ובפורמט .docx'}`);
  }

  let markdown = '';
  const sections: DocSection[] = [];
  let styledHtml = html;

  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');

    // המרת פסקאות שמתחילות ב-# לכותרות, והסרת ה-# מכותרות קיימות
    doc.querySelectorAll('p, h1, h2, h3, h4, h5, h6').forEach(el => {
      const htmlElement = el as HTMLElement;
      const rawText = htmlElement.textContent || '';
      const text = rawText.trim().replace(/\xa0/g, ' ').replace(/[\u200B-\u200D\uFEFF]/g, '');
      
      const match = text.match(/^(#{1,6})\s*(.*?)\s*(?:#*\s*)?$/);
      if (match) {
        const level = match[1].length;
        
        // Strip leading hashes recursively from the first text node in this element to preserve HTML formatting
        const firstTextNode = findFirstTextNode(htmlElement);
        if (firstTextNode) {
          const nodeText = firstTextNode.nodeValue || '';
          const normalized = nodeText.replace(/\xa0/g, ' ').replace(/[\u200B-\u200D\uFEFF]/g, '');
          const stripped = normalized.replace(/^\s*(?:#{1,6})\s*/, '');
          firstTextNode.nodeValue = stripped;
        }

        // Strip trailing hashes recursively from the last text node if the original text ended with #
        if (text.endsWith('#')) {
          const lastTextNode = findLastTextNode(htmlElement);
          if (lastTextNode) {
            const nodeText = lastTextNode.nodeValue || '';
            const normalized = nodeText.replace(/\xa0/g, ' ').replace(/[\u200B-\u200D\uFEFF]/g, '');
            const stripped = normalized.replace(/\s*#+\s*$/, '');
            lastTextNode.nodeValue = stripped;
          }
        }

        // Convert the element to the correct header tag level while preserving its children (bold, colors, links)
        const targetTagName = `h${level}`;
        if (htmlElement.tagName.toLowerCase() !== targetTagName) {
          const heading = doc.createElement(targetTagName);
          while (htmlElement.firstChild) {
            heading.appendChild(htmlElement.firstChild);
          }
          htmlElement.parentNode?.replaceChild(heading, htmlElement);
        }
      }
    });

    // יצירת Markdown לאחר המרת הכותרות, לפני הוספת עיצובים נוספים
    const turndownService = new TurndownService({ headingStyle: 'atx' });
    turndownService.use(gfm);
    markdown = turndownService.turndown(doc.body.innerHTML);

    const headings = doc.querySelectorAll('h1, h2, h3, h4, h5, h6');

    headings.forEach((heading, index) => {
      const id = `sec-${index}`;
      heading.setAttribute('id', id);
      
      let textContent = heading.textContent || 'ללא כותרת';
      // Clean up any remaining leading or trailing # characters and spaces, since they are only for formatting/layout
      textContent = textContent.replace(/^[\s#]+/, '').replace(/[\s#]+$/, '').trim();
      
      let tag = '';

      // זיהוי תגיות טקסט לעיצובים מיוחדים
      if (textContent.includes('[אזהרה]')) {
        tag = 'warning';
        textContent = textContent.replace('[אזהרה]', '').trim();
      } else if (textContent.includes('[מידע]')) {
        tag = 'info';
        textContent = textContent.replace('[מידע]', '').trim();
      } else if (textContent.includes('[הצלחה]')) {
        tag = 'success';
        textContent = textContent.replace('[הצלחה]', '').trim();
      }

      // עדכון הטקסט ב-DOM כדי להסתיר את התגית מהתצוגה - רק אם יש תגית, נשנה את textContent כדי לא להרוס עיצובי פנים סתם
      if (tag) {
        heading.textContent = textContent;
      }

      // הגדרת עיצוב בסיסי
      const baseClasses = 'mt-8 mb-4 font-bold py-3 px-4 rounded-lg border-r-4 shadow-sm scroll-mt-20 block w-full text-right';
      
      if (tag === 'warning') {
        heading.className = `${baseClasses} text-red-900 bg-red-50 border-red-500`;
      } else if (tag === 'info') {
        heading.className = `${baseClasses} text-teal-900 bg-teal-50 border-teal-500`;
      } else if (tag === 'success') {
        heading.className = `${baseClasses} text-green-900 bg-green-50 border-green-500`;
      } else {
        // Theme configurations
        const themes: Record<ThemeColor, { h1: string, h2: string, h3: string, link: string }> = {
          activePresenter: {
            h1: 'text-gray-900 bg-gray-100 border-gray-600',
            h2: 'text-blue-800 bg-blue-50 border-blue-500',
            h3: 'text-emerald-800 bg-emerald-50 border-emerald-500',
            link: 'text-orange-600 hover:text-orange-800'
          },
          classic: {
            h1: 'text-gray-900 bg-gray-100 border-gray-500',
            h2: 'text-yellow-800 bg-yellow-50 border-yellow-400',
            h3: 'text-red-800 bg-red-50 border-red-400',
            link: 'text-blue-500 hover:text-blue-700'
          },
          grayscale: {
            h1: 'text-black bg-gray-200 border-black',
            h2: 'text-gray-900 bg-gray-100 border-gray-600',
            h3: 'text-gray-800 bg-gray-50 border-gray-400',
            link: 'text-gray-600 hover:text-gray-800'
          },
          blue: {
            h1: 'text-blue-900 bg-blue-100 border-blue-800',
            h2: 'text-blue-800 bg-blue-50 border-blue-600',
            h3: 'text-indigo-800 bg-indigo-50 border-indigo-400',
            link: 'text-blue-600 hover:text-blue-800'
          },
          blueGreen: {
            h1: 'text-slate-900 bg-blue-100 border-slate-800',
            h2: 'text-blue-800 bg-blue-50 border-blue-500',
            h3: 'text-emerald-800 bg-emerald-50 border-emerald-500',
            link: 'text-sky-600 hover:text-sky-800'
          },
          green: {
            h1: 'text-emerald-900 bg-emerald-100 border-emerald-800',
            h2: 'text-green-800 bg-green-50 border-green-600',
            h3: 'text-teal-800 bg-teal-50 border-teal-500',
            link: 'text-emerald-600 hover:text-emerald-800'
          },
          greenYellow: {
            h1: 'text-lime-900 bg-lime-100 border-lime-800',
            h2: 'text-green-800 bg-green-50 border-green-600',
            h3: 'text-yellow-800 bg-yellow-50 border-yellow-500',
            link: 'text-lime-600 hover:text-lime-800'
          },
          red: {
            h1: 'text-red-900 bg-red-100 border-red-800',
            h2: 'text-orange-900 bg-orange-50 border-orange-600',
            h3: 'text-amber-900 bg-amber-50 border-amber-600',
            link: 'text-red-600 hover:text-red-800'
          },
          redViolet: {
            h1: 'text-fuchsia-900 bg-fuchsia-100 border-fuchsia-800',
            h2: 'text-purple-900 bg-purple-50 border-purple-600',
            h3: 'text-pink-900 bg-pink-50 border-pink-500',
            link: 'text-fuchsia-600 hover:text-fuchsia-800'
          },
          yellow: {
            h1: 'text-amber-900 bg-amber-100 border-amber-800',
            h2: 'text-yellow-800 bg-yellow-50 border-yellow-500',
            h3: 'text-orange-800 bg-orange-50 border-orange-500',
            link: 'text-amber-600 hover:text-amber-800'
          },
          yellowOrange: {
            h1: 'text-orange-900 bg-orange-100 border-orange-800',
            h2: 'text-amber-900 bg-amber-50 border-amber-600',
            h3: 'text-yellow-900 bg-yellow-50 border-yellow-600',
            link: 'text-orange-600 hover:text-orange-800'
          }
        };

        const activeTheme = themes[theme] || themes.activePresenter;

        if (heading.tagName === 'H2') {
           heading.className = `${baseClasses} ${activeTheme.h2}`;
        } else if (heading.tagName === 'H3') {
           heading.className = `${baseClasses} ${activeTheme.h3}`;
        } else {
           heading.className = `${baseClasses} ${activeTheme.h1}`;
        }
      }

      // התאמות גודל טקסט לפי רמת הכותרת
      if (heading.tagName === 'H1') heading.classList.add('text-3xl');
      if (heading.tagName === 'H2') heading.classList.add('text-xl');
      if (heading.tagName === 'H3') heading.classList.add('text-lg');
      
      const levelChar = heading.tagName[1];
      sections.push({
        id,
        title: textContent,
        level: parseInt(levelChar) || 1,
      });
    });

    // Style standard elements generated by mammoth
    doc.querySelectorAll('p').forEach(p => p.className = 'mb-4 leading-relaxed text-gray-700 text-right');
    doc.querySelectorAll('ul').forEach(ul => ul.className = 'list-disc list-inside mb-4 space-y-1 text-right');
    doc.querySelectorAll('ol').forEach(ol => ol.className = 'list-decimal list-inside mb-4 space-y-1 text-right');
    doc.querySelectorAll('a').forEach(a => a.className = 'text-blue-600 hover:underline');
    doc.querySelectorAll('table').forEach(table => {
      const wrapper = doc.createElement('div');
      wrapper.className = 'overflow-x-auto mb-8 bg-white shadow-sm border border-gray-200 rounded-lg';
      table.parentNode?.insertBefore(wrapper, table);
      wrapper.appendChild(table);
      table.className = 'min-w-full text-right border-collapse border border-gray-200';
    });

    doc.querySelectorAll('tr').forEach(tr => {
      tr.className = 'hover:bg-gray-50 transition-colors group';
    });

    doc.querySelectorAll('th').forEach(th => {
      th.className = 'px-4 py-3 text-sm font-semibold text-gray-900 bg-gray-50 border border-gray-300 text-right';
    });

    doc.querySelectorAll('td').forEach(td => {
      td.className = 'px-4 py-3 text-sm text-gray-700 border border-gray-200 align-top text-right';
    });

    styledHtml = doc.body.innerHTML;
  } catch (domError: any) {
    console.error('DOM parsing or styling failed, using original mammoth HTML:', domError);
    // Graceful fallback to raw mammoth conversion
    const fallbackService = new TurndownService({ headingStyle: 'atx' });
    fallbackService.use(gfm);
    markdown = fallbackService.turndown(html);
    styledHtml = html;

    // Generate basic sections from original h1-h4
    try {
      const fallbackParser = new DOMParser();
      const fallbackDoc = fallbackParser.parseFromString(html, 'text/html');
      fallbackDoc.querySelectorAll('h1, h2, h3, h4, h5, h6').forEach((heading, index) => {
        sections.push({
          id: `sec-${index}`,
          title: heading.textContent || 'ללא כותרת',
          level: parseInt(heading.tagName[1]) || 1
        });
      });
    } catch (e) {
      console.error('Failed to parse fallback sections:', e);
    }
  }

  return {
    html: styledHtml,
    markdown,
    sections,
  };
}
