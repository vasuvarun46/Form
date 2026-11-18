import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { createSpContext } from './lib/spContext';
import Eform from './webparts/eform/components/Eform';
import '../assets/main.css';

document.body.classList.add('gq-eform-fullpage');

async function bootstrap(): Promise<void> {
  const context = await createSpContext();

  let flowUrl = '';
  let approvalFlowUrl = '';

  try {
    const siteUrl = context.pageContext.web.absoluteUrl;
    const res = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('WorkflowConfiguration')/items?$select=Title,Url,ConfigValue,IsActive&$top=200`,
      null
    );

    if (res.ok) {
      const json = await res.json();
      const items = Array.isArray(json?.value)
        ? (json.value as Array<{ Title?: unknown; Url?: unknown; ConfigValue?: unknown; IsActive?: unknown }>)
        : [];

      for (const item of items) {
        const key = String(item?.Title || '').trim().toLowerCase();
        const isActive = item?.IsActive === undefined ? true : !!item.IsActive;
        let urlValue = '';
        if (typeof item.Url === 'string') urlValue = item.Url.trim();
        else if (item.Url && typeof (item.Url as { Url?: string }).Url === 'string')
          urlValue = (item.Url as { Url: string }).Url.trim();
        else if (typeof item.ConfigValue === 'string') urlValue = item.ConfigValue.trim();

        if (!isActive || !urlValue) continue;
        if (key === 'flowurl') flowUrl = urlValue;
        if (key === 'approvalflowurl') approvalFlowUrl = urlValue;
      }
    }
  } catch { /* ignore */ }

  const root = document.getElementById('root');
  if (!root) throw new Error('Root element not found');

  ReactDOM.render(
    <Eform
      description=""
      isDarkTheme={false}
      environmentMessage=""
      hasTeamsContext={false}
      userDisplayName={context.pageContext.user.displayName}
      context={context}
      flowUrl={flowUrl}
      ApprovalFlowUrl={approvalFlowUrl}
    />,
    root
  );
}

bootstrap().catch(console.error);
