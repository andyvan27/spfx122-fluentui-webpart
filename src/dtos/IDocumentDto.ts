export interface IDocumentDto {
  // Strongly typed core fields your UI depends on
  id: number;
  name: string;
  url: string;
  modified: Date;
  modifiedBy: string;
  fileSize: number;
  fileType: string;

  // Dynamic bag for everything else
  fields?: Record<string, unknown>;
}
