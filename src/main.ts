import { enableProdMode } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';
import { environment } from './environments/environment';

const officeJsLoadFailedErrorMessage = `[INIT] OfficeJS load failed`;

Office.onReady()
  .then(({ host, platform }) => {
    if (host != null && platform != null) {
      // We are in an Office host (Outlook, Desktop (PC) / Mac / Online / Mobile),
      console.log(`[INIT] OfficeJS load complete: ${host} on ${platform}`);
      // console.log(`[INIT] Switch on OfficeJS extended error logging`);
      // OfficeExtension.config.extendedErrorLogging = true; // NOT AVAILABLE IN OUTLOOK
    }
    else {
      throw new Error(officeJsLoadFailedErrorMessage);
    }
  })
  .then(() => {
    const isInDialog: boolean = Office && Office.context && Office.context.mailbox == null || false;

    if (isInDialog) {
      console.log(`In dialog - sending message to parent`);
      Office.context.ui.messageParent(JSON.stringify({ type: 'MessageFromDialog', data: 'Hello from dialog' }));
    }
    else {
      if (environment.production) {
        enableProdMode();
      }

      platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    }
  })
  .catch((error) => {
    if (error.message === officeJsLoadFailedErrorMessage) {
      console.warn(officeJsLoadFailedErrorMessage);
    } else {
      console.error(error);
    }
  });
