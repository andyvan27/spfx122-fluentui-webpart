export interface IGroupDto {
  id: string;
  displayName: string;
  mail?: string;
  groupTypes?: string[];
  fields?: Record<string, unknown>;
}
