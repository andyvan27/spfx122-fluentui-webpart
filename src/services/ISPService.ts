import { IDocumentDto } from "../dtos/IDocumentDto";
import { IFieldInfoDto } from "../dtos/IFieldInfoDto";

export interface ISPService {
  /**
   * Loads list fields (columns) for dynamic column generation.
   */
  getListFields(listTitle: string, viewName?: string): Promise<IFieldInfoDto[]>
  
  /**
   * Loads documents from a SharePoint document library and maps them
   * into the strongly-typed IDocument model with a dynamic field bag.
   */
  getDocuments(listTitle: string): Promise<IDocumentDto[]>;

  /**
   * Loads documents from a SharePoint document library in a paged manner.
   */
  getDocumentsPaged(
    listTitle: string,
    viewFieldNames: string[],
    pageSize?: number
  ): AsyncGenerator<IDocumentDto[], void, unknown>;

  /**
   * Loads list items from a SharePoint list in a paged manner with server-side paging, sorting and filtering using CAML based.
   */
  getDocumentsStreamPaged(
    listTitle: string,
    viewXml: string,
  ): AsyncGenerator<IDocumentDto[], void, unknown>
}
