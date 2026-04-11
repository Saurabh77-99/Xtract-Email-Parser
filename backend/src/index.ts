import { Hono } from "hono";
import { z } from "zod";
import { zValidator } from "@hono/zod-validator";
import { db } from "./db/index.js";
import { rules, messages, results } from "./db/schema.js";
import { eq, sql, inArray, and } from "drizzle-orm";

const app = new Hono();

app.get("/", (c) => c.text("Email Parser API running"));

app.get("/healthz", async (c) => {
  try {
    await db.select({ count: sql<number>`count(*)` }).from(rules);
    return c.json({
      db: "ok",
      url: process.env.TURSO_URL ? "set" : "MISSING",
      token: process.env.TURSO_AUTH_TOKEN ? "set" : "MISSING",
    });
  } catch (e: any) {
    return c.json({ db: "error", message: e.message }, 500);
  }
});

// ============================================================
// RULES
// ============================================================

app.get("/rules", async (c) => {
  const allRules = await db.select().from(rules);
  return c.json(allRules);
});

const createRuleSchema = z.object({
  name: z.string(),
  criteriaQuery: z.string(),
  targetFields: z.string(),
  extractionMode: z.string().optional(),
});

app.post("/rules", zValidator("json", createRuleSchema), async (c) => {
  const { name, criteriaQuery, targetFields, extractionMode } =
    c.req.valid("json");
  try {
    JSON.parse(targetFields);
  } catch (e) {
    return c.json({ error: "Invalid JSON in target_fields" }, 400);
  }
  await db.insert(rules).values({
    name,
    criteriaQuery,
    targetFields,
    extractionMode: extractionMode || "regex",
    isActive: true,
  });
  return c.json({ status: "success" });
});

app.delete("/rules/:id", async (c) => {
  const id = parseInt(c.req.param("id"));
  try {
    const ruleMessages = await db
      .select({ messageId: messages.messageId })
      .from(messages)
      .where(eq(messages.ruleId, id));
    const messageIds = ruleMessages.map((m) => m.messageId);
    if (messageIds.length > 0) {
      await db.delete(results).where(inArray(results.messageId, messageIds));
      await db.delete(messages).where(eq(messages.ruleId, id));
    }
    await db.delete(rules).where(eq(rules.id, id));
    return c.json({ status: "deleted" });
  } catch (err: any) {
    return c.json({ error: err.message }, 500);
  }
});

// ============================================================
// BATCH INGEST — processes multiple emails in one request
// ============================================================

const batchIngestSchema = z.object({
  ruleId: z.number(),
  messages: z.array(
    z.object({
      messageId: z.string(),
      subject: z.string(),
      sender: z.string(),
      date: z.string(),
      body: z.string(),
    }),
  ),
});

app.post("/batch-ingest", zValidator("json", batchIngestSchema), async (c) => {
  const { ruleId, messages: msgs } = c.req.valid("json");

  const ruleRes = await db.query.rules.findFirst({
    where: eq(rules.id, ruleId),
  });
  if (!ruleRes) return c.json({ error: "Rule not found" }, 404);

  let targetFields: Record<string, any> = {};
  try {
    targetFields = JSON.parse(ruleRes.targetFields);
  } catch {
    return c.json({ error: "Invalid rule format" }, 500);
  }

  let successCount = 0;

  for (const msg of msgs) {
    try {
      // Insert message (skip if already exists)
      await db
        .insert(messages)
        .values({
          messageId: msg.messageId,
          ruleId,
          subject: msg.subject,
          sender: msg.sender,
          rawBody: msg.body,
        })
        .onConflictDoNothing();

      // Extract fields using existing logic
      const extractedData: Array<{ key: string; value: string }> = [];

      for (const [key, fieldDef] of Object.entries(targetFields)) {
        let matchedValue = "";

        if (typeof fieldDef === "string") {
          try {
            const regex = new RegExp(fieldDef, "i");
            const match = msg.body.match(regex);
            if (match) matchedValue = match[1] || match[0];
          } catch {}
        } else if (typeof fieldDef === "object" && fieldDef !== null) {
          const def = fieldDef as { type: string; anchor: string; end: string };
          matchedValue = extractTextAnchor(
            msg.body,
            def.anchor,
            def.type,
            def.end,
          );
        }

        if (matchedValue) extractedData.push({ key, value: matchedValue });
      }

      for (const data of extractedData) {
        await db
          .insert(results)
          .values({
            messageId: msg.messageId,
            ruleId,
            key: data.key,
            value: data.value,
          })
          .onConflictDoNothing();
      }

      successCount++;
    } catch (err: any) {
      console.error(
        "Batch ingest error on " + msg.messageId + ": " + err.message,
      );
    }
  }

  return c.json({ successCount });
});

app.patch("/rules/:id/toggle", async (c) => {
  const id = parseInt(c.req.param("id"));
  const rule = await db.query.rules.findFirst({ where: eq(rules.id, id) });
  if (!rule) return c.json({ error: "Rule not found" }, 404);
  await db
    .update(rules)
    .set({ isActive: !rule.isActive })
    .where(eq(rules.id, id));
  return c.json({ status: "success", isActive: !rule.isActive });
});

// ============================================================
// STATS + ACTIVITY
// ============================================================

app.get("/stats", async (c) => {
  const [msgCount] = await db
    .select({ count: sql<number>`count(*)` })
    .from(messages);
  const [resCount] = await db
    .select({ count: sql<number>`count(*)` })
    .from(results);
  const lastActivity = await db
    .select()
    .from(messages)
    .orderBy(sql`created_at DESC`)
    .limit(1);
  return c.json({
    totalMessages: msgCount.count,
    totalExtractions: resCount.count,
    lastSync: lastActivity[0]?.createdAt || "Never",
  });
});

app.get("/activity", async (c) => {
  const recent = await db
    .select({
      messageId: messages.messageId,
      subject: messages.subject,
      sender: messages.sender,
      createdAt: messages.createdAt,
    })
    .from(messages)
    .orderBy(sql`created_at DESC`)
    .limit(10);
  return c.json(recent);
});

// ============================================================
// INGEST — simple body
// ============================================================

const ingestSchema = z.object({
  ruleId: z.number(),
  messageId: z.string(),
  subject: z.string(),
  sender: z.string(),
  rawBody: z.string(),
});

app.post("/ingest", zValidator("json", ingestSchema), async (c) => {
  const { ruleId, messageId, subject, sender, rawBody } = c.req.valid("json");

  const ruleRes = await db.query.rules.findFirst({
    where: eq(rules.id, ruleId),
  });
  if (!ruleRes) return c.json({ error: "Rule not found" }, 404);

  await db
    .insert(messages)
    .values({ messageId, ruleId, subject, sender, rawBody })
    .onConflictDoNothing();

  let targetFields: Record<string, any> = {};
  try {
    targetFields = JSON.parse(ruleRes.targetFields);
  } catch {
    return c.json({ error: "Invalid rule format" }, 500);
  }

  const extractedData: Array<{ key: string; value: string }> = [];

  for (const [key, fieldDef] of Object.entries(targetFields)) {
    let matchedValue = "";

    if (typeof fieldDef === "string") {
      // Legacy regex mode
      try {
        const regex = new RegExp(fieldDef, "i");
        const match = rawBody.match(regex);
        if (match) matchedValue = match[1] || match[0];
      } catch {}
    } else if (typeof fieldDef === "object" && fieldDef !== null) {
      // New text anchor mode
      const def = fieldDef as { type: string; anchor: string; end: string };
      matchedValue = extractTextAnchor(rawBody, def.anchor, def.type, def.end);
    }

    if (matchedValue) extractedData.push({ key, value: matchedValue });
  }

  for (const data of extractedData) {
    await db
      .insert(results)
      .values({ messageId, ruleId, key: data.key, value: data.value })
      .onConflictDoNothing();
  }

  return c.json({ status: "success", extracted: extractedData.length });
});

// ============================================================
// INGEST RICH — body + attachments
// ============================================================

const ingestRichSchema = z.object({
  ruleId: z.number(),
  messageId: z.string(),
  subject: z.string(),
  sender: z.string(),
  rawBody: z.string(),
  attachments: z
    .array(
      z.object({
        name: z.string(),
        mimeType: z.string(),
        data: z.string(),
        isExcel: z.boolean().optional(),
        extractedText: z.string().optional(),
      }),
    )
    .optional(),
});

app.post("/ingest-rich", zValidator("json", ingestRichSchema), async (c) => {
  const { ruleId, messageId, subject, sender, rawBody, attachments } =
    c.req.valid("json");

  const ruleRes = await db.query.rules.findFirst({
    where: eq(rules.id, ruleId),
  });
  if (!ruleRes) return c.json({ error: "Rule not found" }, 404);

  await db
    .insert(messages)
    .values({ messageId, ruleId, subject, sender, rawBody })
    .onConflictDoNothing();

  let targetFields: Record<string, any> = {};
  try {
    targetFields = JSON.parse(ruleRes.targetFields);
  } catch {}

  const extractedData: Array<{ key: string; value: string }> = [];

  // Extract from body
  for (const [key, fieldDef] of Object.entries(targetFields)) {
    let matchedValue = "";
    if (typeof fieldDef === "string") {
      try {
        const regex = new RegExp(fieldDef, "i");
        const match = rawBody.match(regex);
        if (match) matchedValue = match[1] || match[0];
      } catch {}
    } else if (typeof fieldDef === "object" && fieldDef !== null) {
      const def = fieldDef as { type: string; anchor: string; end: string };
      matchedValue = extractTextAnchor(rawBody, def.anchor, def.type, def.end);
    }
    if (matchedValue) extractedData.push({ key, value: matchedValue });
  }

  // Store attachment metadata
  if (attachments && attachments.length > 0) {
    for (const att of attachments) {
      extractedData.push({
        key: "attachment_" + att.name,
        value: att.mimeType,
      });
      if (att.extractedText && att.extractedText.length > 0) {
        extractedData.push({
          key: "content_" + att.name,
          value: att.extractedText.substring(0, 2000),
        });
      }
    }
  }

  for (const data of extractedData) {
    await db
      .insert(results)
      .values({ messageId, ruleId, key: data.key, value: data.value })
      .onConflictDoNothing();
  }

  return c.json({ status: "success", extracted: extractedData.length });
});

// ============================================================
// INGEST DIRECT — for contact builder, no rule needed
// ============================================================

const ingestContactSchema = z.object({
  messageId: z.string(),
  subject: z.string(),
  sender: z.string(),
  senderName: z.string(),
  senderDomain: z.string(),
  date: z.string(),
});

app.post(
  "/ingest-contact",
  zValidator("json", ingestContactSchema),
  async (c) => {
    const contact = c.req.valid("json");
    // Store as a special rule-0 contact entry
    return c.json({ status: "success", contact });
  },
);

// ============================================================
// SUMMARY + EXPORT
// ============================================================

app.get("/summary/:ruleId", async (c) => {
  const ruleId = parseInt(c.req.param("ruleId"));
  const sender = c.req.query("sender");
  const from = c.req.query("from");
  const to = c.req.query("to");

  const allRows = await db
    .select({
      messageId: messages.messageId,
      subject: messages.subject,
      sender: messages.sender,
      createdAt: messages.createdAt,
      key: results.key,
      value: results.value,
    })
    .from(messages)
    .leftJoin(
      results,
      and(
        eq(results.messageId, messages.messageId),
        eq(results.ruleId, messages.ruleId),
      ),
    )
    .where(eq(messages.ruleId, ruleId));

  let filtered = sender
    ? allRows.filter((r) =>
        r.sender?.toLowerCase().includes(sender.toLowerCase()),
      )
    : allRows;
  if (from)
    filtered = filtered.filter((r) => r.createdAt && r.createdAt >= from);
  if (to) filtered = filtered.filter((r) => r.createdAt && r.createdAt <= to);

  const grouped: Record<string, any> = {};
  filtered.forEach((row) => {
    if (!grouped[row.messageId]) {
      grouped[row.messageId] = {
        messageId: row.messageId,
        subject: row.subject,
        sender: row.sender,
        createdAt: row.createdAt,
        fields: {},
      };
    }
    if (row.key) grouped[row.messageId].fields[row.key] = row.value;
  });

  return c.json({
    rule: ruleId,
    totalEmails: Object.keys(grouped).length,
    generatedAt: new Date().toISOString(),
    data: Object.values(grouped),
  });
});

app.get("/export/:ruleId", async (c) => {
  const ruleId = parseInt(c.req.param("ruleId"));
  const allResults = await db
    .select({
      messageId: messages.messageId,
      subject: messages.subject,
      sender: messages.sender,
      key: results.key,
      value: results.value,
      createdAt: results.createdAt,
    })
    .from(results)
    .innerJoin(
      messages,
      and(
        eq(results.messageId, messages.messageId),
        eq(results.ruleId, messages.ruleId),
      ),
    )
    .where(eq(messages.ruleId, ruleId));

  if (allResults.length === 0)
    return c.text("No data found for this rule.", 404);

  const grouped: Record<string, any> = {};
  const keys = new Set<string>();
  for (const row of allResults) {
    if (!grouped[row.messageId]) {
      grouped[row.messageId] = {
        messageId: row.messageId,
        subject: row.subject,
        sender: row.sender,
        createdAt: row.createdAt,
      };
    }
    grouped[row.messageId][row.key] = row.value;
    keys.add(row.key);
  }

  const header = [
    "messageId",
    "subject",
    "sender",
    "createdAt",
    ...Array.from(keys),
  ];
  const csvRows = [header.join(",")];
  for (const rowId in grouped) {
    const row = grouped[rowId];
    csvRows.push(
      header
        .map((k) => `"${(row[k] || "").toString().replace(/"/g, '""')}"`)
        .join(","),
    );
  }

  c.header("Content-Type", "text/csv");
  c.header(
    "Content-Disposition",
    `attachment; filename="export-rule-${ruleId}.csv"`,
  );
  return c.text(csvRows.join("\n"));
});

// ============================================================
// BROWSE
// ============================================================

app.get("/browse/:ruleId", async (c) => {
  const ruleId = parseInt(c.req.param("ruleId"));
  const data = await db
    .select({
      messageId: messages.messageId,
      subject: messages.subject,
      key: results.key,
      value: results.value,
      createdAt: messages.createdAt,
    })
    .from(results)
    .innerJoin(
      messages,
      and(
        eq(results.messageId, messages.messageId),
        eq(results.ruleId, messages.ruleId),
      ),
    )
    .where(eq(messages.ruleId, ruleId))
    .orderBy(sql`messages.created_at DESC`);

  const grouped: Record<string, any> = {};
  data.forEach((row) => {
    if (!grouped[row.messageId]) {
      grouped[row.messageId] = {
        messageId: row.messageId,
        subject: row.subject,
        createdAt: row.createdAt,
        data: {},
      };
    }
    grouped[row.messageId].data[row.key] = row.value;
  });
  return c.json(Object.values(grouped));
});

// ============================================================
// TEXT ANCHOR EXTRACTION (server-side)
// ============================================================

function extractTextAnchor(
  body: string,
  anchor: string,
  type: string,
  boundary: string,
): string {
  if (!anchor || !body) return "";
  const idx = body.indexOf(anchor);
  if (idx === -1) return "";

  if (type === "after") {
    const start = idx + anchor.length;
    const rest = body.substring(start).trimStart();
    return extractByBoundary(rest, boundary);
  } else if (type === "before") {
    const before = body.substring(0, idx).trimEnd();
    const lines = before.split("\n");
    return extractByBoundary(lines[lines.length - 1].trim(), boundary);
  }
  return "";
}

function extractByBoundary(text: string, boundary: string): string {
  if (boundary === "word") return (text.match(/^\S+/) || [""])[0];
  if (boundary === "line") return text.split("\n")[0].trim();
  if (boundary === "paragraph") return text.split(/\n\n/)[0].trim();
  return text.split("\n")[0].trim();
}

export default app;
