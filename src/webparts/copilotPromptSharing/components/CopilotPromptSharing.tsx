import * as React from 'react';
import Dashboard from './Dashboard/Dashboard';
import { ICopilotPromptSharingProps } from './ICopilotPromptSharingProps';

const CopilotPromptSharing = (
  props: ICopilotPromptSharingProps
): JSX.Element => {
  return <Dashboard context={props.context} />;
};

export default CopilotPromptSharing;