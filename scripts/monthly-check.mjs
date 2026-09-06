// Reads backoffice-data.json (the same file Back-Office syncs to GitHub)
// and figures out which compliance items are due this calendar month
// (or already overdue). Writes a subject/body for the email step to use,
// and sends a phone push notification directly via ntfy.sh.

import fs from "fs";

const DATA_PATH = "backoffice-data.json";

const LABELS = {
  dMedical: "Medical Card",
  dIfta: "IFTA",
  dIrp: "IRP",
  dInspection: "Annual Inspection",
};

function daysUntil(dateStr) {
  const target = new Date(dateStr + "T00:00:00");
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  return Math.round((target - now) / 86400000);
}

function isDueThisMonthOrOverdue(dateStr) {
  const d = new Date(dateStr + "T00:00:00");
  const now = new Date();
  const sameMonth = d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth();
  return sameMonth || d < now;
}

function formatDate(dateStr) {
  const d = new Date(dateStr + "T00:00:00");
  return d.toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" });
}

function writeOutput(name, value) {
  const outPath = process.env.GITHUB_OUTPUT;
  if (!outPath) return;
  const delimiter = `EOF_${name}_${Date.now()}`;
  fs.appendFileSync(outPath, `${name}<<${delimiter}\n${value}\n${delimiter}\n`);
}

let raw = {};
try {
  raw = JSON.parse(fs.readFileSync(DATA_PATH, "utf8"));
} catch (e) {
  console.log("No backoffice-data.json found yet, or it's not valid JSON.");
}

const dates = raw.complianceDates || {};
const monthName = new Date().toLocaleDateString("en-US", { month: "long", year: "numeric" });

const dueItems = [];
for (const [id, label] of Object.entries(LABELS)) {
  const dateStr = dates[id];
  if (dateStr && isDueThisMonthOrOverdue(dateStr)) {
    dueItems.push({ label, date: formatDate(dateStr), overdue: daysUntil(dateStr) < 0 });
  }
}

let subject, body;
if (Object.keys(dates).length === 0) {
  subject = `Fleet Ledger: no compliance data connected yet`;
  body = `This monthly check ran, but no backoffice-data.json was found in the repo. Connect GitHub Sync in Back-Office (Compliance Cluster) so this check has dates to look at.`;
} else if (dueItems.length === 0) {
  subject = `Fleet Ledger: nothing due in ${monthName}`;
  body = `No compliance items (Medical Card, IFTA, IRP, Inspection) are due in ${monthName}. You're clear for now.`;
} else {
  subject = `Fleet Ledger: ${dueItems.length} item${dueItems.length > 1 ? "s" : ""} due in ${monthName}`;
  body =
    `Compliance items due in ${monthName}:\n\n` +
    dueItems.map((i) => `- ${i.label}: ${i.date}${i.overdue ? " (OVERDUE)" : ""}`).join("\n");
}

console.log(subject);
console.log(body);

writeOutput("subject", subject);
writeOutput("body", body);
writeOutput("has_items", dueItems.length > 0 ? "true" : "false");

// ---- Push notification via ntfy.sh (free, no account needed) ----
const ntfyTopic = process.env.NTFY_TOPIC;
if (ntfyTopic) {
  try {
    const res = await fetch(`https://ntfy.sh/${ntfyTopic}`, {
      method: "POST",
      headers: {
        Title: subject,
        Priority: dueItems.some((i) => i.overdue) ? "urgent" : "default",
      },
      body,
    });
    console.log("ntfy push status:", res.status);
  } catch (e) {
    console.error("ntfy push failed:", e.message);
  }
} else {
  console.log("NTFY_TOPIC secret not set - skipping push notification.");
}
