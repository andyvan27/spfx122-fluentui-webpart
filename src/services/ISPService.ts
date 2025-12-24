import { IDocumentDto } from "../dtos/IDocumentDto";
import { IListItemDto } from "../dtos/IListItemDto";
import { IFieldInfoDto } from "../dtos/IFieldInfoDto";

export interface ISPService {
  /**
   * Loads documents from a SharePoint document library and maps them
   * into the strongly-typed IDocument model with a dynamic field bag.
   */
  getDocuments(listTitle: string): Promise<IDocumentDto[]>;

  /**
   * Loads raw list items from a SharePoint list and maps them into
   * the generic IListItem model (id + dynamic field bag).
   */
  getListItems(listTitle: string): Promise<IListItemDto[]>;

  /**
   * Loads list fields (columns) for dynamic column generation.
   */
  getListFields(listTitle: string): Promise<IFieldInfoDto[]>;

  /**
   * Executes a CAML query and returns raw items mapped into IListItem.
   */
  getItemsByCaml(listTitle: string, viewXml: string): Promise<IListItemDto[]>;

  /**
   * Loads items from a specific folder (useful for document libraries).
   */
  getFolderItems(serverRelativeFolderPath: string): Promise<IDocumentDto[]>;
}
