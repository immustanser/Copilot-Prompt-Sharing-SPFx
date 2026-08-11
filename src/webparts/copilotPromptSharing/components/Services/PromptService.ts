import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';
import { getSP } from '../../pnpjsConfig';
import { IPrompt } from '../Models/IPrompt';
import { PermissionKind } from '@pnp/sp/security';
import '@pnp/sp/fields';
import '@pnp/sp/site-users/web';

const LIST_TITLE = 'Stewart AI Prompt Library';

interface ISPListPromptItem {
    Id: number;
    Title: string;
    AiTool?: string;
    PromptText?: string;
    UseCase?: string;
    Tags?: string;
    Status?: string;
    Department?: string;
}

export interface INewPromptItem {
    title: string;
    aiTool: string;
    department: string;
    useCase: string;
    tags: string;
    promptText: string;
}

export interface IDashboardData {
    currentUserId: number;
    prompts: IPrompt[];
    departmentChoices: string[];
    aiToolChoices: string[];
    permissions: {
        canAdd: boolean;
        canEdit: boolean;
        canDelete: boolean;
    };
}

const PROMPT_SELECT_FIELDS = [
    'Id', 'Title', 'AiTool', 'PromptText', 'UseCase', 'Tags', 'Status', 'Department'
] as const;

function mapItem(item: ISPListPromptItem): IPrompt {
    return {
        Id: item.Id,
        Title: item.Title || '',
        AiTool: item.AiTool || '',
        PromptText: item.PromptText || '',
        UseCase: item.UseCase || '',
        Tags: item.Tags || '',
        Status: item.Status || '',
        Department: item.Department || ''
    };
}

function buildVisibilityFilter(currentUserId: number): string {
    return (
        `Status eq 'Approved' or ` +
        `((Status eq 'Send for Approval' or Status eq 'Rejected') and AuthorId eq ${currentUserId})`
    );
}

export default class PromptService {
    private sp: SPFI;

    constructor(context: WebPartContext) {
        this.sp = getSP(context);
    }

    /**
     * Initializes all dashboard data in two network round trips:
     *   Round 1 – fetch current user ID (required for the item visibility filter)
     *   Round 2 – six calls in parallel: items, dept choices, AI tool choices,
     *              and all three permission checks.
     */
    public async initializeDashboard(): Promise<IDashboardData> {
        const list = this.sp.web.lists.getByTitle(LIST_TITLE);

        // Round 1 – only fetch the Id field to minimise payload
        const currentUser = await this.sp.web.currentUser.select('Id')();
        const currentUserId: number = currentUser.Id;

        const visibilityFilter = buildVisibilityFilter(currentUserId);

        // Round 2 – all remaining calls in parallel
        const [
            items,
            deptField,
            aiToolField,
            canAdd,
            canEdit,
            canDelete
        ] = await Promise.all([
            list.items
                .select(...PROMPT_SELECT_FIELDS)
                .filter(visibilityFilter)
                .top(5000)() as Promise<ISPListPromptItem[]>,
            list.fields.getByInternalNameOrTitle('Department')(),
            list.fields.getByInternalNameOrTitle('AiTool')(),
            list.currentUserHasPermissions(PermissionKind.AddListItems),
            list.currentUserHasPermissions(PermissionKind.EditListItems),
            list.currentUserHasPermissions(PermissionKind.DeleteListItems)
        ]);

        return {
            currentUserId,
            prompts: items.map(mapItem),
            departmentChoices: (deptField as { Choices?: string[] }).Choices || [],
            aiToolChoices: (aiToolField as { Choices?: string[] }).Choices || [],
            permissions: { canAdd, canEdit, canDelete }
        };
    }

    /**
     * Lightweight post-CRUD refresh – fetches only list items (no permissions,
     * no field metadata). Pass the userId already obtained at dashboard init.
     */
    public async refreshPrompts(currentUserId: number): Promise<IPrompt[]> {
        const items = await this.sp.web.lists
            .getByTitle(LIST_TITLE)
            .items
            .select(...PROMPT_SELECT_FIELDS)
            .filter(buildVisibilityFilter(currentUserId))
            .top(5000)() as ISPListPromptItem[];

        return items.map(mapItem);
    }

    /** @deprecated Use initializeDashboard() for initial load. Kept for compatibility. */
    public async getPermissions(): Promise<{ canAdd: boolean; canEdit: boolean; canDelete: boolean }> {
        const list = this.sp.web.lists.getByTitle(LIST_TITLE);
        const [canAdd, canEdit, canDelete] = await Promise.all([
            list.currentUserHasPermissions(PermissionKind.AddListItems),
            list.currentUserHasPermissions(PermissionKind.EditListItems),
            list.currentUserHasPermissions(PermissionKind.DeleteListItems)
        ]);
        return { canAdd, canEdit, canDelete };
    }

    /** @deprecated Use initializeDashboard() for initial load. Kept for compatibility. */
    public async getPrompts(): Promise<IPrompt[]> {
        const currentUser = await this.sp.web.currentUser.select('Id')();
        const items = await this.sp.web.lists
            .getByTitle(LIST_TITLE)
            .items
            .select(...PROMPT_SELECT_FIELDS)
            .filter(buildVisibilityFilter(currentUser.Id))
            .top(5000)() as ISPListPromptItem[];
        return items.map(mapItem);
    }

    public async createPrompt(item: INewPromptItem): Promise<void> {
        await this.sp.web.lists
            .getByTitle(LIST_TITLE)
            .items
            .add({
                Title: item.title,
                AiTool: item.aiTool,
                PromptText: item.promptText,
                UseCase: item.useCase,
                Tags: item.tags,
                Department: item.department,
                Status: 'Send for Approval'
            });
    }

    public async updatePrompt(itemId: number, item: INewPromptItem): Promise<void> {
        await this.sp.web.lists
            .getByTitle(LIST_TITLE)
            .items
            .getById(itemId)
            .update({
                Title: item.title,
                AiTool: item.aiTool,
                PromptText: item.promptText,
                UseCase: item.useCase,
                Tags: item.tags,
                Department: item.department,
                Status: 'Send for Approval'
            });
    }

    public async deletePrompt(itemId: number): Promise<void> {
        await this.sp.web.lists
            .getByTitle(LIST_TITLE)
            .items
            .getById(itemId)
            .delete();
    }

    /** @deprecated Use initializeDashboard() for initial load. Kept for compatibility. */
    public async getDepartmentChoices(): Promise<string[]> {
        const field = await this.sp.web.lists
            .getByTitle(LIST_TITLE)
            .fields
            .getByInternalNameOrTitle('Department')();
        return (field as { Choices?: string[] }).Choices || [];
    }

    /** @deprecated Use initializeDashboard() for initial load. Kept for compatibility. */
    public async getAiToolChoices(): Promise<string[]> {
        const field = await this.sp.web.lists
            .getByTitle(LIST_TITLE)
            .fields
            .getByInternalNameOrTitle('AiTool')();
        return (field as { Choices?: string[] }).Choices || [];
    }
}
