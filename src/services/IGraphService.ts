import { IUserDto } from "../dtos/IUserDto";
import { ITeamDto } from "../dtos/ITeamDto";
import { IGroupDto } from "../dtos/IGroupDto";
import { IUserPhotoDto } from "../dtos/IUserPhotoDto";

export interface IGraphService {
  /**
   * Returns the currently logged-in user.
   */
  getCurrentUser(): Promise<IUserDto>;

  /**
   * Returns a user by Azure AD object ID or UPN.
   */
  getUser(userId: string): Promise<IUserDto>;

  /**
   * Returns the manager of the specified user (or current user if omitted).
   * Returns undefined if the user has no manager.
   */
  getManager(userId?: string): Promise<IUserDto | undefined>;

  /**
   * Returns the user's profile photo.
   * Returns undefined if the user has no photo.
   */
  getUserPhoto(userId: string): Promise<IUserPhotoDto | undefined>;

  /**
   * Returns the Teams the current user is a member of.
   */
  getJoinedTeams(): Promise<ITeamDto[]>;

  /**
   * Returns the Microsoft 365 groups the current user belongs to.
   */
  getMemberGroups(): Promise<IGroupDto[]>;
}
