export interface ITeamDto {
  id: string;
  displayName: string;
  description?: string;
  fields?: Record<string, unknown>;
}
