let msalInstance;
let graphAccount = null;
let lastCheckedKey = null; // avoids re-querying for the same set of "To" addresses

const els = {
  status: document.getElementById("status"),
  signin: document.getElementById("signin"),
  result: document.getElementById("result"),
};

function showState(name) {
  ["status", "signin", "result"].forEach((k) => {
    els[k].style.display = k === name ? "block" : "none";
  });
}

Office.onReady(() => {
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: APP_CONFIG.clientId,
      authority: APP_CONFIG.authority,
      redirectUri: APP_CONFIG.redirectUri,
    },
    cache: { cacheLocation: "localStorage" },
  });

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) graphAccount = accounts[0];

  document.getElementById("signin-btn").addEventListener("click", signIn);

  // Check whatever is already in "To" when the pane opens, then watch for changes.
  checkCurrentRecipients();
  Office.context.mailbox.item.to.addHandlerAsync(
    Office.EventType.RecipientsChanged,
    checkCurrentRecipients
  );
});

async function signIn() {
  try {
    const loginResp = await msalInstance.loginPopup({
      scopes: ["Mail.Read"],
    });
    graphAccount = loginResp.account;
    checkCurrentRecipients();
  } catch (err) {
    renderError("Sign-in failed: " + err.message);
  }
}

async function getGraphToken() {
  if (!graphAccount) throw new Error("NO_ACCOUNT");
  try {
    const resp = await msalInstance.acquireTokenSilent({
      scopes: ["Mail.Read"],
      account: graphAccount,
    });
    return resp.accessToken;
  } catch (err) {
    const resp = await msalInstance.acquireTokenPopup({ scopes: ["Mail.Read"] });
    return resp.accessToken;
  }
}

function checkCurrentRecipients() {
  Office.context.mailbox.item.to.getAsync(async (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) return;

    const recipients = (result.value || []).filter((r) => r.emailAddress);
    if (recipients.length === 0) {
      lastCheckedKey = null;
      showState("status");
      els.status.innerHTML = "<p>Waiting for a recipient in the \"To\" field&hellip;</p>";
      return;
    }

    const key = recipients.map((r) => r.emailAddress.toLowerCase()).sort().join(",");
    if (key === lastCheckedKey) return; // nothing meaningful changed
    lastCheckedKey = key;

    if (!graphAccount) {
      showState("signin");
      return;
    }

    showState("status");
    els.status.innerHTML = "<p>Checking Sent Items&hellip;</p>";

    try {
      const token = await getGraphToken();
      const findings = [];
      for (const r of recipients) {
        const hit = await findRecentSend(token, r.emailAddress);
        findings.push({ address: r.emailAddress, hit });
      }
      renderResults(findings);
    } catch (err) {
      renderError("Couldn't check Sent Items: " + err.message);
    }
  });
}

// Looks in Sent Items (last 7 days) for a message sent to this address
// where it was the sole recipient (no other To/Cc) — i.e. a genuine 1:1.
async function findRecentSend(token, emailAddress) {
  const since = new Date();
  since.setDate(since.getDate() - 7);
  const sinceIso = since.toISOString();

  const url =
    "https://graph.microsoft.com/v1.0/me/mailFolders/SentItems/messages" +
    "?$filter=" + encodeURIComponent(`sentDateTime ge ${sinceIso}`) +
    "&$select=subject,sentDateTime,toRecipients,ccRecipients" +
    "&$orderby=sentDateTime desc&$top=100";

  const resp = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!resp.ok) throw new Error(`Graph error ${resp.status}`);
  const data = await resp.json();

  const target = emailAddress.toLowerCase();
  const matches = (data.value || []).filter((msg) => {
    const to = msg.toRecipients || [];
    const cc = msg.ccRecipients || [];
    const toAddresses = to.map((p) => p.emailAddress.address.toLowerCase());
    const isSoleRecipient = to.length === 1 && cc.length === 0;
    return isSoleRecipient && toAddresses.includes(target);
  });

  return matches.length > 0 ? matches[0] : null;
}

function renderResults(findings) {
  showState("result");
  const parts = findings.map(({ address, hit }) => {
    if (!hit) {
      return `<div class="banner banner-clear">No recent contact — ${address}</div>`;
    }
    const sentDate = new Date(hit.sentDateTime);
    const when = sentDate.toLocaleDateString(undefined, {
      weekday: "short",
      month: "short",
      day: "numeric",
    });
    return `
      <div class="banner banner-hit">Already emailed ${address} this week</div>
      <div class="card">
        <p class="subject">${escapeHtml(hit.subject || "(no subject)")}</p>
        <p class="meta">Sent ${when}</p>
      </div>
    `;
  });
  els.result.innerHTML = parts.join("");
}

function renderError(message) {
  showState("result");
  els.result.innerHTML = `<div class="error">${escapeHtml(message)}</div>`;
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}
