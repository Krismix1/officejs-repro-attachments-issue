/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

function getAttachments(item: typeof Office.context.mailbox.item) {
  // getAttachmentContentAsync() is only supported above 1.8
  if (!Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
    return [];
  }

  const attachments = item.attachments.filter(
    (attachment) => attachment.attachmentType === "file" && attachment.isInline === true
  );
  console.log("Filtered attachments:", attachments.length);

  if (!attachments) {
    return [];
  }

  return Promise.all(
    attachments.map(
      (attachment) =>
        new Promise((resolve) => {
          item.getAttachmentContentAsync(attachment.id, (result) => resolve(result));
        })
    )
  );
}

async function runOnce() {
  console.log("Trying to fetch attachments...", new Date());
  const item = Office.context.mailbox.item;
  try {
    const attachments = await getAttachments(item);
    console.log(attachments);
  } catch (e) {
    console.log("----------------------------------------------- never executed", e);
  }
}

export async function run() {
  await runOnce();
  // 10 minutes, longer than token expiration
  // weird...here it does not trigger the error
  setTimeout(() => runOnce().then(() => console.log("timeout finished")), 10 * 60 * 1000);
}
