Office.onReady(() => {
  document.getElementById("sendButton").onclick = handleSubmit;
});

async function handleSubmit() {
  try {
    const item = Office.context.mailbox.item;
    const messageContent = document.getElementById("messageContent").value;

    // Subject
    let subject = "";
    if (typeof item.subject === "string") {
      subject = item.subject; // read mode
    } else if (item.subject && item.subject.getAsync) {
      subject = await new Promise((resolve) =>
        item.subject.getAsync((res) => resolve(res.value || ""))
      );
    }

    // From
    let from = "";
    if (item.from && item.from.emailAddress) {
      from = item.from.emailAddress; // read mode
    } else {
      from = Office.context.mailbox.userProfile.emailAddress; // compose mode
    }

    // To
    let to = [];
    if (Array.isArray(item.to)) {
      to = item.to.map((r) => r.emailAddress); // read mode
    } else if (item.to && item.to.getAsync) {
      to = await new Promise((resolve) =>
        item.to.getAsync((res) =>
          resolve((res.value || []).map((r) => r.emailAddress))
        )
      );
    }

    // Body (plain text)
    const bodyPlainText = await new Promise((resolve) =>
      item.body.getAsync(Office.CoercionType.Text, (res) =>
        resolve(res.value || "")
      )
    );

    const payload = {
      itemId: item.itemId || "",
      subject,
      from,
      to,
      date:
        item.dateTimeCreated instanceof Date
          ? item.dateTimeCreated.toISOString()
          : new Date().toISOString(),
      bodyPlainText,
      input: messageContent,
    };

    // Send to your n8n webhook
    const webhookUrl =
      "https://bannister.app.n8n.cloud/webhook-test/479e6f18-8845-488e-9f97-96900a30037a";

    await fetch(webhookUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    alert("Sent to n8n ✅");
    document.getElementById("messageContent").value = "";
  } catch (err) {
    console.error(err);
    alert("Failed to send to n8n ❌: " + (err.message || err));
  }
}
