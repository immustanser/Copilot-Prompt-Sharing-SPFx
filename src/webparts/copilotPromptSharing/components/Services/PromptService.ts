import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';
import { getSP } from '../../pnpjsConfig';
import { IPrompt } from '../Models/IPrompt';
import { PermissionKind } from '@pnp/sp/security';
import '@pnp/sp/fields';

interface ISPListPromptItem {
    Id: number;
    Title: string;
    AIToolOfChoice?: string;
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

export default class PromptService {
    private sp: SPFI;

    constructor(context: WebPartContext) {
        this.sp = getSP(context);
    }

    public async getPermissions() {
        const list = this.sp.web.lists.getByTitle(
            'Copilot Prompt Sharing'
        );

        return {
            canAdd:
                await list.currentUserHasPermissions(
                    PermissionKind.AddListItems
                ),

            canEdit:
                await list.currentUserHasPermissions(
                    PermissionKind.EditListItems
                ),

            canDelete:
                await list.currentUserHasPermissions(
                    PermissionKind.DeleteListItems
                )
        };
    }

    public async getPrompts(): Promise<IPrompt[]> {
        const items: ISPListPromptItem[] = await this.sp.web.lists
            .getByTitle('Copilot Prompt Sharing')
            .items
            .select(
                'Id',
                'Title',
                'AIToolOfChoice',
                'PromptText',
                'UseCase',
                'Tags',
                'Status',
                'Department'
            )
            .top(5000)();

        return items.map((item: ISPListPromptItem): IPrompt => {
            return {
                Id: item.Id,
                Title: item.Title || '',
                AIToolOfChoice: item.AIToolOfChoice || '',
                PromptText: item.PromptText || '',
                UseCase: item.UseCase || '',
                Tags: item.Tags || '',
                Status: item.Status || '',
                Department: item.Department || ''
            };
        });
    }

    public async createPrompt(item: INewPromptItem): Promise<void> {
        await this.sp.web.lists
            .getByTitle('Copilot Prompt Sharing')
            .items
            .add({
                Title: item.title,
                AIToolOfChoice: item.aiTool,
                PromptText: item.promptText,
                UseCase: item.useCase,
                Tags: item.tags,
                Department: item.department,
                Status: 'Send for Approval'
            });
    }

    public async updatePrompt(
        itemId: number,
        item: INewPromptItem
    ): Promise<void> {
        await this.sp.web.lists
            .getByTitle('Copilot Prompt Sharing')
            .items
            .getById(itemId)
            .update({
                Title: item.title,
                AIToolOfChoice: item.aiTool,
                PromptText: item.promptText,
                UseCase: item.useCase,
                Tags: item.tags,
                Department: item.department,
                Status: 'Send for Approval'
            });
    }

    public async deletePrompt(
        itemId: number
    ): Promise<void> {

        await this.sp.web.lists
            .getByTitle('Copilot Prompt Sharing')
            .items
            .getById(itemId)
            .delete();
    }

    public async getDepartmentChoices(): Promise<string[]> {

        const field = await this.sp.web.lists
            .getByTitle('Copilot Prompt Sharing')
            .fields
            .getByInternalNameOrTitle('Department')();

        return field.Choices || [];
    }
}
