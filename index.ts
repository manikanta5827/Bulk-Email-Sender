import { Hono } from "hono";
import * as XLSX from "xlsx";
import { Resend } from "resend";

const app = new Hono();

const COMPANY_VARIANTS = new Set([
  "company name", "company_name", "companyname", "company",
  "organization", "organisation", "org", "business", "business name",
  "client", "account", "customer",
]);

const EMAIL_VARIANTS = new Set([
  "email", "e-mail", "email address", "email_address", "emailaddress",
  "mail", "e mail", "contact", "email id", "emailid",
]);

function normalizeHeader(raw: string): string {
  return raw
    .toLowerCase()
    .replace(/[_\-\s]+/g, " ")
    .trim();
}

function findMatchingColumn(
  headers: string[],
  variants: Set<string>
): string | null {
  for (const h of headers) {
    const norm = normalizeHeader(h);
    if (variants.has(norm)) return h;
  }
  for (const h of headers) {
    const norm = normalizeHeader(h);
    for (const v of variants) {
      if (norm.includes(v) || v.includes(norm)) return h;
    }
  }
  return null;
}

function replacePlaceholders(
  text: string,
  data: { company_name: string; email: string }
): string {
  return text
    .replace(/\{\{company_name\}\}/gi, data.company_name)
    .replace(/\$\{company_name\}/gi, data.company_name)
    .replace(/\{\{company\s*name\}\}/gi, data.company_name)
    .replace(/\$\{company\s*name\}/gi, data.company_name)
    .replace(/\{\{email\}\}/gi, data.email)
    .replace(/\$\{email\}/gi, data.email);
}

app.get("/", async (c) => {
  const html = await Bun.file("./public/index.html").text();
  return c.html(html);
});

app.post("/api/upload", async (c) => {
  try {
    const formData = await c.req.formData();
    const file = formData.get("file");

    if (!file || !(file instanceof File)) {
      return c.json({ error: "No file provided" }, 400);
    }

    const MAX_FILE_SIZE = 10 * 1024 * 1024;
    if (file.size > MAX_FILE_SIZE) {
      return c.json({ error: "File too large. Maximum size is 10 MB." }, 400);
    }

    const buffer = await file.arrayBuffer();
    const data = new Uint8Array(buffer);

    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) {
      return c.json({ error: "No sheets found in file" }, 400);
    }

    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      return c.json({ error: "Failed to read sheet: " + sheetName }, 400);
    }
    const rows = XLSX.utils.sheet_to_json<Record<string, string>>(sheet, {
      defval: "",
    });

    if (rows.length === 0) {
      return c.json({ error: "No data rows found in file" }, 400);
    }

    const headers = Object.keys(rows[0]!);
    const companyCol = findMatchingColumn(headers, COMPANY_VARIANTS);
    const emailCol = findMatchingColumn(headers, EMAIL_VARIANTS);

    if (!emailCol) {
      return c.json({
        error:
          "Could not find an email column. Expected header like: Email, E-mail, Mail, etc. Found: " +
          headers.join(", "),
      }, 400);
    }

    const recipients = rows
      .map((row) => ({
        company_name: (companyCol ? row[companyCol] : "") || "",
        email: (row[emailCol] || "").trim(),
      }))
      .filter((r) => r.email && r.email.includes("@"));

    if (recipients.length === 0) {
      return c.json({ error: "No valid email addresses found" }, 400);
    }

    const capped = recipients.length > 100;
    const result = capped ? recipients.slice(0, 100) : recipients;

    return c.json({
      recipients: result,
      headers,
      companyCol,
      emailCol,
      total: result.length,
      capped,
      message: capped
        ? "Free tier limit: 100 emails/day. Only the first 100 recipients will be used."
        : undefined,
    });
  } catch (err) {
    return c.json(
      { error: "Failed to parse file: " + (err as Error).message },
      400
    );
  }
});

app.post("/api/send", async (c) => {
  let body: {
    from: string;
    subject: string;
    html: string;
    recipients: { company_name: string; email: string }[];
  };

  try {
    body = await c.req.json();
  } catch {
    return c.json({ error: "Invalid JSON body" }, 400);
  }

  const { from, subject, html, recipients } = body;

  const resendApiKey = process.env.RESEND_API_KEY;
  if (!resendApiKey) {
    return c.json({ error: "RESEND_API_KEY environment variable is not set" }, 500);
  }
  if (!from) {
    return c.json({ error: "Sender email (from) is required" }, 400);
  }
  if (!recipients?.length) {
    return c.json({ error: "No recipients provided" }, 400);
  }
  if (recipients.length > 100) {
    return c.json({ error: "Free tier limit: maximum 100 emails per day. You have " + recipients.length + " recipients." }, 400);
  }

  const encoder = new TextEncoder();

  const stream = new ReadableStream({
    async start(controller) {
      const enqueue = (data: Record<string, unknown>) => {
        controller.enqueue(
          encoder.encode("data: " + JSON.stringify(data) + "\n\n")
        );
      };

      const resend = new Resend(resendApiKey);

      let sent = 0;
      let failed = 0;
      const total = recipients.length;

      enqueue({ event: "start", total });

      const BATCH_SIZE = 4;
      const BATCH_DELAY_MS = 1000;

      for (let i = 0; i < recipients.length; i += BATCH_SIZE) {
        if (controller.desiredSize === null) {
          return;
        }

        const batch = recipients.slice(i, i + BATCH_SIZE);

        const results = await Promise.allSettled(
          batch.map(async (recipient) => {
            const emailSubject = replacePlaceholders(subject, recipient);
            const emailHtml = replacePlaceholders(html, recipient);

            const { error } = await resend.emails.send({
              from,
              to: [recipient.email],
              subject: emailSubject,
              html: emailHtml,
            });

            if (error) {
              throw new Error(error.message);
            }

            return recipient;
          })
        );

        for (let j = 0; j < results.length; j++) {
          const result = results[j]!;
          const recipient = batch[j]!;

          if (result.status === "fulfilled") {
            sent++;
            enqueue({
              event: "progress",
              sent,
              failed,
              total,
              current: recipient.email,
              company: recipient.company_name,
              status: "sent",
            });
          } else {
            failed++;
            enqueue({
              event: "error",
              sent,
              failed,
              total,
              email: recipient.email,
              company: recipient.company_name,
              error: (result as PromiseRejectedResult).reason?.message ?? "Unknown error",
              status: "failed",
            });
          }
        }

        if (i + BATCH_SIZE < recipients.length) {
          await new Promise((r) => setTimeout(r, BATCH_DELAY_MS));
        }
      }

      enqueue({ event: "complete", sent, failed, total });
      controller.close();
    },
  });

  return new Response(stream, {
    headers: {
      "Content-Type": "text/event-stream",
      "Cache-Control": "no-cache",
      Connection: "keep-alive",
    },
  });
});

Bun.serve({
  fetch: app.fetch,
  port: 3000,
});

console.log("Bulk Email Sender running at http://localhost:3000");
