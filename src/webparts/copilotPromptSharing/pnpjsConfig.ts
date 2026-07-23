import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFI, SPFx } from '@pnp/sp';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

let _sp: SPFI | undefined;

export const getSP = (context?: WebPartContext): SPFI => {
  if (!_sp && context) {
    _sp = spfi().using(SPFx(context));
  }

  return _sp as SPFI;
};
