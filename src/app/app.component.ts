import { Component } from '@angular/core';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

  openDialog() {
    const { host, protocol } = window.location;
    const url = `https://auth-test.leap.services/oauth/authorize?client_id=S639FGYFEDPLHTWM&redirect_uri=https://outlookapp-dist-leapaws.s3-ap-southeast-2.amazonaws.com/index.html&scope=*&response_type=token&newsession&view_target=outlook_android&isInDialog=true`;
    const options = { width: 30, height: 30, displayInIframe: false };
    //const options = { width: 30, height: 30, displayInIframe: false, promptBeforeOpen: true },
    //const options = { width: 30, height: 30, displayInIframe: false, promptBeforeOpen: false },

    console.log(url);

    Office.context.ui.displayDialogAsync(
      url,
      options,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {

          const dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived,
            (msg) => {
              console.log(`Office.EventType.DialogMessageReceived handler!`)
              if (msg) {
                console.log(msg);
                console.log('MESSAGE:', JSON.parse(msg.message));
                console.log(`Message received - closing dialog from parent`);
                dialog.close();
              }
            });

          dialog.addEventHandler(Office.EventType.DialogEventReceived,
            (evt) => {
              console.log(`Office.EventType.DialogEventReceived handler!`)
              if (evt) { console.log(evt) }
            });
        }
        else {
          // Oh no
          console.log(`ERROR SHOWING DIALOG`);
          console.log(result.error);
          console.log(result.diagnostics);
        }
      });
  }
}
