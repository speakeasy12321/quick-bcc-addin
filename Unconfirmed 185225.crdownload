/* Quick BCC Add-in – taskpane.js */

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    loadSavedEmail();

    document.getElementById("add-bcc-btn").addEventListener("click", addBcc);
    document.getElementById("save-btn").addEventListener("click", saveEmail);
  }
});

// ── Roaming settings key ──────────────────────────────────────────────────────
const SETTINGS_KEY = "quickBccEmail";

// ── Load the saved email from roaming settings ────────────────────────────────
function loadSavedEmail() {
  const settings = Office.context.roamingSettings;
  const email = settings.get(SETTINGS_KEY);

  const display = document.getElementById("saved-email-display");
  const btn     = document.getElementById("add-bcc-btn");

  if (email) {
    display.textContent = email;
    btn.disabled = false;
  } else {
    display.textContent = "Not configured yet";
    btn.disabled = true;
  }
}

// ── Save a new email to roaming settings ──────────────────────────────────────
function saveEmail() {
  const input = document.getElementById("email-input").value.trim();

  if (!isValidEmail(input)) {
    showStatus("Please enter a valid email address.", "error");
    return;
  }

  const settings = Office.context.roamingSettings;
  settings.set(SETTINGS_KEY, input);

  settings.saveAsync(function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      document.getElementById("email-input").value = "";
      loadSavedEmail();
      showStatus("Contact saved!", "success");
    } else {
      showStatus("Could not save. Please try again.", "error");
    }
  });
}

// ── Add the saved email as BCC on the current message ────────────────────────
function addBcc() {
  const settings = Office.context.roamingSettings;
  const email    = settings.get(SETTINGS_KEY);

  if (!email) {
    showStatus("No contact saved yet.", "error");
    return;
  }

  const item = Office.context.mailbox.item;

  item.bcc.addAsync(
    [{ emailAddress: email }],
    function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        showStatus("✓ Added " + email + " as BCC!", "success");
      } else {
        showStatus("Failed to add BCC: " + result.error.message, "error");
      }
    }
  );
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function showStatus(message, type) {
  const el = document.getElementById("status");
  el.textContent = message;
  el.className = type;
  setTimeout(() => { el.textContent = ""; el.className = ""; }, 4000);
}
