const crypto = require("crypto");

let cachedToken = null;
let tokenExpiresAt = 0;

class PayPalError extends Error {
  constructor(message, status, details) {
    super(message);
    this.name = "PayPalError";
    this.status = status;
    this.details = details;
  }
}

function config() {
  const clientId = process.env.PAYPAL_CLIENT_ID;
  const clientSecret = process.env.PAYPAL_CLIENT_SECRET;
  const mode = (process.env.PAYPAL_MODE || "sandbox").toLowerCase();
  const defaultBaseUrl = mode === "live"
    ? "https://api-m.paypal.com"
    : "https://api-m.sandbox.paypal.com";
  const baseUrl = (process.env.PAYPAL_BASE_URL || defaultBaseUrl).replace(/\/$/, "");

  if (!clientId || !clientSecret) {
    throw new Error("PAYPAL_CLIENT_ID and PAYPAL_CLIENT_SECRET must be configured.");
  }
  if (!["https://api-m.sandbox.paypal.com", "https://api-m.paypal.com"].includes(baseUrl)) {
    throw new Error("PAYPAL_BASE_URL must be an official PayPal API URL.");
  }
  if (mode === "sandbox" && !baseUrl.includes("sandbox")) {
    throw new Error("PAYPAL_MODE=sandbox cannot use the live PayPal API URL.");
  }

  return { clientId, clientSecret, mode, baseUrl };
}

async function parseResponse(response) {
  const text = await response.text();
  if (!text) return {};
  try {
    return JSON.parse(text);
  } catch {
    return { message: text };
  }
}

async function getAccessToken() {
  if (cachedToken && Date.now() < tokenExpiresAt) return cachedToken;

  const { clientId, clientSecret, baseUrl } = config();
  const response = await fetch(`${baseUrl}/v1/oauth2/token`, {
    method: "POST",
    headers: {
      Authorization: `Basic ${Buffer.from(`${clientId}:${clientSecret}`).toString("base64")}`,
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: "grant_type=client_credentials",
    signal: AbortSignal.timeout(15000),
  });
  const body = await parseResponse(response);
  if (!response.ok || !body.access_token) {
    throw new PayPalError(body.error_description || "Unable to authenticate with PayPal.", response.status, body);
  }

  cachedToken = body.access_token;
  tokenExpiresAt = Date.now() + Math.max(0, Number(body.expires_in || 300) - 60) * 1000;
  return cachedToken;
}

async function paypalRequest(path, options = {}) {
  const { baseUrl } = config();
  const accessToken = await getAccessToken();
  const response = await fetch(`${baseUrl}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
    signal: AbortSignal.timeout(20000),
  });
  const body = await parseResponse(response);
  if (!response.ok) {
    const issue = body.details?.[0]?.issue;
    const description = body.details?.[0]?.description;
    throw new PayPalError(description || body.message || "PayPal request failed.", response.status, {
      name: body.name,
      issue,
      debugId: response.headers.get("paypal-debug-id"),
    });
  }
  return body;
}

async function createPayout({ senderBatchId, senderItemId, recipientEmail, amount, currency, invoiceId }) {
  return paypalRequest("/v1/payments/payouts", {
    method: "POST",
    headers: { "PayPal-Request-Id": crypto.randomUUID() },
    body: JSON.stringify({
      sender_batch_header: {
        sender_batch_id: senderBatchId,
        recipient_type: "EMAIL",
        email_subject: "Supplier payment sent",
        email_message: `Payment for invoice ${invoiceId} has been sent.`,
      },
      items: [{
        recipient_type: "EMAIL",
        recipient_wallet: "PAYPAL",
        receiver: recipientEmail,
        amount: { value: Number(amount).toFixed(2), currency },
        note: `Supplier invoice ${invoiceId}`,
        sender_item_id: senderItemId,
      }],
    }),
  });
}

function getPayoutBatch(batchId) {
  return paypalRequest(`/v1/payments/payouts/${encodeURIComponent(batchId)}?page=1&page_size=100&total_required=true`);
}

function verifyWebhook(headers, webhookEvent) {
  const webhookId = process.env.PAYPAL_WEBHOOK_ID;
  if (!webhookId) throw new Error("PAYPAL_WEBHOOK_ID must be configured before accepting webhooks.");
  return paypalRequest("/v1/notifications/verify-webhook-signature", {
    method: "POST",
    body: JSON.stringify({
      auth_algo: headers["paypal-auth-algo"],
      cert_url: headers["paypal-cert-url"],
      transmission_id: headers["paypal-transmission-id"],
      transmission_sig: headers["paypal-transmission-sig"],
      transmission_time: headers["paypal-transmission-time"],
      webhook_id: webhookId,
      webhook_event: webhookEvent,
    }),
  });
}

module.exports = { PayPalError, config, createPayout, getPayoutBatch, verifyWebhook };
