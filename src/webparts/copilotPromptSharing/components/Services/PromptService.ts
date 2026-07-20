import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';
import { getSP } from '../../pnpjsConfig';
import { IPrompt } from '../Models/IPrompt';

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

export default class PromptService {
  private sp: SPFI;

  constructor(context: WebPartContext) {
    this.sp = getSP(context);
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
}