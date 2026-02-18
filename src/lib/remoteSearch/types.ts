export type DebatifyApiTagHit = {
  tag: string;
  citation: string;
  markdown: string;
};

export type DebatifyTagHit = {
  id: string;
  tag: string;
  citation: string;
  richHtml: string;
  plainText: string;
  copyText: string;
  paragraphXml: string[];
  sourcePath: string;
};
