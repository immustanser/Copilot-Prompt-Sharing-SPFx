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
    isApprover: boolean;
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

/**
 * Builds the OData filter that controls which list items the current user sees.
 *
 * Visibility rules:
 *  - Approved        → everyone
 *  - Send for Approval → creator always; approvers see ALL pending items
 *  - Rejected        → creator only (approvers do NOT see others' rejections)
 */
function buildVisibilityFilter(currentUserId: number, isApprover: boolean): string {
    if (isApprover) {
        return (
            `Status eq 'Approved' or ` +
            `Status eq 'Send for Approval' or ` +
            `(Status eq 'Rejected' and AuthorId eq ${currentUserId})`
        );
    }
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
     * Initializes all dashboard data in three network round trips:
     *   Round 1 – current user Id (needed by every subsequent filter)
     *   Round 2 – approver check + field choices + permission checks in parallel
     *             (approver result is required before the items query can be built)
     *   Round 3 – list items with the role-correct visibility filter
     *
     * Total HTTP calls: 1 + 5 + 1 = 7 (same count as the previous implementation).
     */
    public async initializeDashboard(): Promise<IDashboardData> {
        const list = this.sp.web.lists.getByTitle(LIST_TITLE);

        // ── Round 1: resolve current user ────────────────────────────────────────
        const currentUser = await this.sp.web.currentUser.select('Id')();
        const currentUserId: number = currentUser.Id;

        // ── Round 2: approver check + field metadata + permissions ────────────────
        // The approver result must be known before the items query can be built,
        // so items are intentionally moved to Round 3.
        const [
            approverItems,
            deptField,
            aiToolField,
            canAdd,
            canEdit,
            canDelete
        ] = await Promise.all([
            // Person fields expose their numeric Id as <InternalName>Id in OData.
            this.sp.web.lists
                .getByTitle('Copilot Prompt Approvers')
                .items
                .select('Id')
                .filter(`ApproverId eq ${currentUserId}`)
                .top(1)() as Promise<{ Id: number }[]>,
            list.fields.getByInternalNameOrTitle('Department')(),
            list.fields.getByInternalNameOrTitle('AiTool')(),
            list.currentUserHasPermissions(PermissionKind.AddListItems),
            list.currentUserHasPermissions(PermissionKind.EditListItems),
            list.currentUserHasPermissions(PermissionKind.DeleteListItems)
        ]);

        const isApprover = approverItems.length > 0;

        // ── Round 3: list items with the role-correct security filter ─────────────
        const items = await list.items
            .select(...PROMPT_SELECT_FIELDS)
            .filter(buildVisibilityFilter(currentUserId, isApprover))
            .top(5000)() as ISPListPromptItem[];

        return {
            currentUserId,
            isApprover,
            prompts: items.map(mapItem),
            departmentChoices: (deptField as { Choices?: string[] }).Choices || [],
            aiToolChoices: (aiToolField as { Choices?: string[] }).Choices || [],
            permissions: { canAdd, canEdit, canDelete }
        };
    }

    /**
     * Lightweight post-CRUD refresh – fetches only list items (no permissions,
     * no field metadata). Pass the userId and isApprover flag cached at init.
     */
    public async refreshPrompts(currentUserId: number, isApprover: boolean): Promise<IPrompt[]> {
        const items = await this.sp.web.lists
            .getByTitle(LIST_TITLE)
            .items
            .select(...PROMPT_SELECT_FIELDS)
            .filter(buildVisibilityFilter(currentUserId, isApprover))
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
            .filter(buildVisibilityFilter(currentUser.Id, false))
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

    public async approvePrompt(itemId: number): Promise<void> {
        await this.sp.web.lists
            .getByTitle(LIST_TITLE)
            .items
            .getById(itemId)
            .update({ Status: 'Approved' });
    }

    public async rejectPrompt(itemId: number): Promise<void> {
        await this.sp.web.lists
            .getByTitle(LIST_TITLE)
            .items
            .getById(itemId)
            .update({ Status: 'Rejected' });
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
