export interface IUserDto {
  /**
   * Strongly typed core fields from Microsoft Graph /me or /users/{id}.
   */
  id: string;
  displayName: string;
  mail?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  mobilePhone?: string;
  userPrincipalName: string;

  /**
   * Dynamic bag for any additional fields returned by Graph.
   * This allows you to access arbitrary profile properties,
   * extension attributes, or beta fields without updating the interface.
   */
  fields?: Record<string, unknown>;
}
