export interface IListItemDto {
  /**
   * Strongly typed core fields that every SharePoint list item has.
   */
  id: number;

  /**
   * Optional title field (common in many lists, but not guaranteed).
   */
  title?: string;

  /**
   * Dynamic bag containing all raw fields returned by SharePoint.
   * This allows the UI or service layer to access any field
   * without needing to update the interface.
   */
  fields: Record<string, unknown>;
}
