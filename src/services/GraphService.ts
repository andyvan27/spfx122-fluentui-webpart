import { graphfi, SPFx as GraphSPFx } from "@pnp/graph";

// PnPjs Graph module augmentation
import "@pnp/graph/users";
import "@pnp/graph/users/photos";
import "@pnp/graph/groups";
import "@pnp/graph/teams";

import { IGraphService } from "./IGraphService";
import { IUserDto } from "../dtos/IUserDto";
import { IUserPhotoDto } from "../dtos/IUserPhotoDto";
import { IGroupDto } from "../dtos/IGroupDto";
import { ITeamDto } from "../dtos/ITeamDto";

import { WebPartContext } from "@microsoft/sp-webpart-base";

export class GraphService implements IGraphService {
    private graph;

    constructor(context: WebPartContext) {
        this.graph = graphfi().using(GraphSPFx(context));
    }

    // -------------------------------------------------------
    // CURRENT USER
    // -------------------------------------------------------
    public async getCurrentUser(): Promise<IUserDto> {
        const u = await this.graph.me();

        return {
            id: u.id ?? "",
            displayName: u.displayName ?? "",
            mail: u.mail ?? undefined,
            userPrincipalName: u.userPrincipalName ?? "",
            jobTitle: u.jobTitle ?? undefined,
            department: u.department ?? undefined,
            officeLocation: u.officeLocation ?? undefined,
            mobilePhone: u.mobilePhone ?? undefined,
            fields: { ...u }
        };
    }

    // -------------------------------------------------------
    // USER BY ID
    // -------------------------------------------------------
    public async getUser(userId: string): Promise<IUserDto> {
        const u = await this.graph.users.getById(userId)();

        return {
            id: u.id ?? "",
            displayName: u.displayName ?? "",
            mail: u.mail ?? undefined,
            userPrincipalName: u.userPrincipalName ?? "",
            jobTitle: u.jobTitle ?? undefined,
            department: u.department ?? undefined,
            officeLocation: u.officeLocation ?? undefined,
            mobilePhone: u.mobilePhone ?? undefined,
            fields: { ...u }
        };
    }

    // -------------------------------------------------------
    // MANAGER
    // -------------------------------------------------------
    public async getManager(userId?: string): Promise<IUserDto | undefined> {
        try {
            const manager = userId
                ? await this.graph.users.getById(userId).manager()
                : await this.graph.me.manager();

            if (!manager) return undefined;

            return {
                id: manager.id ?? "",
                displayName: manager.displayName ?? "",
                mail: manager.mail ?? undefined,
                userPrincipalName: manager.userPrincipalName ?? "",
                jobTitle: manager.jobTitle ?? undefined,
                department: manager.department ?? undefined,
                officeLocation: manager.officeLocation ?? undefined,
                mobilePhone: manager.mobilePhone ?? undefined,
                fields: { ...manager }
            };
        } catch {
            return undefined;
        }
    }

    // -------------------------------------------------------
    // USER PHOTO
    // -------------------------------------------------------
    public async getUserPhoto(userId: string): Promise<IUserPhotoDto | undefined> {
        try {
            // Force TS to treat this as a Graph user endpoint, not a DTO
            const userEndpoint = this.graph.users.getById(userId) as any;

            const blob = await userEndpoint.photo.getBlob();

            return {
                userId,
                blob,
                url: URL.createObjectURL(blob)
            };
        } catch {
            return undefined;
        }
    }

    // -------------------------------------------------------
    // JOINED TEAMS
    // -------------------------------------------------------
    public async getJoinedTeams(): Promise<ITeamDto[]> {
        const teams = await this.graph.me.joinedTeams();

        return teams.map(t => ({
            id: t.id ?? "",
            displayName: t.displayName ?? "",
            description: t.description ?? undefined,
            fields: { ...t }
        }));
    }

    // -------------------------------------------------------
    // MEMBER GROUPS
    // -------------------------------------------------------
    public async getMemberGroups(): Promise<IGroupDto[]> {
        const objects = await this.graph.me.memberOf();

        // Filter only Microsoft 365 Groups
        const groups = objects.filter((o: any) =>
            o["@odata.type"] === "#microsoft.graph.group"
        );

        return groups.map((g: any) => ({
            id: g.id ?? "",
            displayName: g.displayName ?? "",
            mail: g.mail ?? undefined,
            fields: { ...g }
        }));
    }
}
