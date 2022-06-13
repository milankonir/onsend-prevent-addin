import { ContextReplacementPlugin } from "webpack";

async function onSend(event: Office.AddinCommands.Event) {
  await Office.onReady();

  event.completed({ allowEvent: false });
}

Object.assign(globalThis, {
  onSend,
});

Office.onReady();
