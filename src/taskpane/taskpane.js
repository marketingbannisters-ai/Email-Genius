Office.onReady(() => {
  const btn = document.getElementById("sendButton");
  if (btn) btn.onclick = handleSubmit;
});

async function handleSubmit() {
  const btn = document.getElementById("sendButton");
  const loader = document.getElementById("loader");

  try {
    startLoading(btn, loader);

    const item = Office.context.mailbox.item;
    const note = (document.getElementById("messageContent")?.value ?? "").toString();
    const messageContent = document.getElementById("messageContent").value;

    // Subject
    let subject = "";
    if (typeof item.subject === "string") {
      subject = item.subject;
    } else if (item.subject && item.subject.getAsync) {
      subject = await new Promise((resolve) =>
        item.subject.getAsync((res) => resolve(res.value || ""))
      );
    }

    // From
    let from = "";
    if (item.from && item.from.emailAddress) {
      from = item.from.emailAddress;
    } else {
      from = Office.context.mailbox.userProfile.emailAddress;
    }

    // To
    let to = [];
    if (Array.isArray(item.to)) {
      to = item.to.map((r) => r.emailAddress);
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

    // 1) Call n8n
    const resp = await fetch("https://bannister.app.n8n.cloud/webhook/ee70b84d-ff6f-4b41-80ba-57e8ab0f4a35", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ payload }),
    });
    if (!resp.ok) {
      const errText = await resp.text().catch(() => "");
      throw new Error(`Webhook failed: ${resp.status} ${errText}`.trim());
    }

    // 2) Parse flexibly
    let replyHtml = "";
    let raw = await resp.text();
    try {
      const data = JSON.parse(raw);
      replyHtml = data.replyHtml
        ? String(data.replyHtml)
        : (data.replyText ? `<div>${String(data.replyText).replace(/\n/g,"<br>")}</div>` : "");
    } catch {
      replyHtml = `<div>${raw.replace(/\n/g,"<br>")}</div>`;
    }

    // Handle accidental double-encoded JSON string
    if (/^"\s*</.test(replyHtml) && />\s*"$/.test(replyHtml)) {
      try { replyHtml = JSON.parse(replyHtml); } catch {}
    }
    if (!replyHtml || !replyHtml.trim()) throw new Error("No replyHtml/replyText in webhook response");

    // 3) Sanitize + cap
    replyHtml = sanitizeReplyHtml(replyHtml);
    const MAX = 200000;
    if (replyHtml.length > MAX) replyHtml = replyHtml.slice(0, MAX) + "<p>…(truncated)</p>";

    replyHtml = enforceFontSizeDom(replyHtml, "13pt"); 

    // 4) Draft it
    const isReadSurface = typeof item.displayReplyForm === "function";
    if (isReadSurface) {
      try {
        item.displayReplyForm(replyHtml);
        showInfo("Reply opened (HTML) ✅");
        return;
      } catch {
        item.displayReplyForm(stripToPlainText(replyHtml));
        showInfo("Reply opened (plain text) ✅");
        return;
      }
    }

    await delay(300);
    try {
      await setBody(item, replyHtml, "Html");
      await saveDraft(item);
      showInfo("Draft saved (HTML) ✅");
    } catch (errHtml) {
      await setBody(item, stripToPlainText(replyHtml), "Text");
      await saveDraft(item);
      showInfo("Draft saved (plain text) ✅");
    }

    const el = document.getElementById("messageContent");
    if (el) el.value = "";

  } catch (err) {
    console.error(err);
    showError(`Failed to create draft: ${err?.message || err}`);
  } finally {
    stopLoading(btn, loader);
  }
}

/* --- loading helpers --- */
function startLoading(btn, loader) {
  if (btn) {
    btn.disabled = true;
    btn.dataset._label = btn.innerText;
    btn.innerText = "Generating…";
    btn.setAttribute("aria-busy", "true");
  }
  if (loader) {
    loader.classList.remove("hidden");
    loader.setAttribute("aria-hidden", "false");
  }
}
function stopLoading(btn, loader) {
  if (btn) {
    btn.disabled = false;
    btn.innerText = btn.dataset._label || "Create Draft";
    btn.removeAttribute("aria-busy");
  }
  if (loader) {
    loader.classList.add("hidden");
    loader.setAttribute("aria-hidden", "true");
  }
}

/* ---------- your existing helpers below ---------- */
function delay(ms){return new Promise(r=>setTimeout(r,ms));}
function setBody(item, content, coercion /* "Html" | "Text" */) {
  return new Promise((resolve, reject) => {
    if (!item?.body?.setAsync) return reject(new Error("Compose body API unavailable"));
    const type = coercion === "Text" ? Office.CoercionType.Text : Office.CoercionType.Html;
    item.body.setAsync(content, { coercionType: type }, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) return resolve();
      reject(res.error || new Error("Load failed"));
    });
  });
}
function saveDraft(item){return new Promise((resolve)=>{ if(!item?.saveAsync) return resolve(); item.saveAsync(()=>resolve()); });}
function stripToPlainText(html){ const tmp=document.createElement("div"); tmp.innerHTML=html; return tmp.textContent||tmp.innerText||""; }
function describeOfficeError(err){ if(!err) return ""; const parts=[err.code, err.name, err.message||err.description].filter(Boolean); return parts.join(" | "); }
function sanitizeReplyHtml(html){
  if(!html) return "<p></p>";
  return html
    .replace(/<!doctype[\s\S]*?>/gi,"")
    .replace(/<\/?html[\s\S]*?>/gi,"")
    .replace(/<\/?head[\s\S]*?>/gi,"")
    .replace(/<\/?body[\s\S]*?>/gi,"")
    .replace(/<script[\s\S]*?<\/script>/gi,"")
    .replace(/<style[\s\S]*?<\/style>/gi,"")
    .replace(/<link[^>]*rel=["']?stylesheet["']?[^>]*>/gi,"")
    .replace(/<\/?iframe[\s\S]*?>/gi,"")
    .replace(/\son\w+="[^"]*"/gi,"")
    .replace(/\son\w+='[^']*'/gi,"")
    .replace(/src="http:\/\//gi,'src="https://');
}
function showInfo(msg){ const it=Office.context.mailbox.item; it.notificationMessages.replaceAsync("n8nStatus",{type:Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,message:msg,icon:"Icon.16x16",persistent:false}); }
function showError(msg){ const it=Office.context.mailbox.item; it.notificationMessages.replaceAsync("n8nError",{type:Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,message:msg}); }
function enforceFontSizeDom(html, pt = "13pt") {
  if (!html) return "";
  const container = document.createElement("div");
  container.innerHTML = html;

  const TAGS = new Set(["P","DIV","SPAN","LI","TD","TH","A","STRONG","EM","B","I","U"]);
  const hasFontSize = (style) => /(^|;)\s*font-size\s*:/i.test(style);

  // Add font-size to elements that don't already specify one
  const walker = document.createTreeWalker(container, NodeFilter.SHOW_ELEMENT, null);
  let node;
  while ((node = walker.nextNode())) {
    if (!TAGS.has(node.tagName)) continue;
    const style = node.getAttribute("style") || "";
    if (!hasFontSize(style)) {
      node.setAttribute(
        "style",
        (style ? style.replace(/\s*$/,"; ") : "") + `font-size:${pt}; mso-bidi-font-size:${pt};`
      );
    }
  }

  // Wrap stray text nodes (text sitting directly inside container or blocks)
  function wrapLooseText(parent) {
    const nodes = Array.from(parent.childNodes);
    for (const n of nodes) {
      if (n.nodeType === Node.TEXT_NODE && n.nodeValue.trim()) {
        const span = document.createElement("span");
        span.setAttribute("style", `font-size:${pt}; mso-bidi-font-size:${pt};`);
        span.textContent = n.nodeValue;
        parent.replaceChild(span, n);
      } else if (n.nodeType === Node.ELEMENT_NODE) {
        wrapLooseText(n);
      }
    }
  }
  wrapLooseText(container);

  return container.innerHTML;
}
