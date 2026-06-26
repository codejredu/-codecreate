export interface DocSection {
  id: string;
  title: string;
  level: number;
}

export interface DocData {
  html: string;
  markdown: string;
  sections: DocSection[];
}
