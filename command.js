function sendToTodoist(event) {
  const item = Office.context.mailbox.item;

  const itemId = item.itemId;
  const encodedId = encodeURIComponent(itemId);

  const link = `https://outlook.office.com/mail/deeplink/read/${encodedId}?ItemID=${encodedId}&exvsurl=1`;

  const message = {
    toRecipients: ["add.task.re4ph2xcffcnfnwk@todoist.net"],
    subject: "Email link to Todoist",
    htmlBody: `Odkaz na e-mail:<br><br><a href="${link}">${link}</a>`
  };

  Office.context.mailbox.displayNewMessageForm(message);

  event.completed();
}
