// Cloudflare Pages Function — lives at /api/upload once deployed.
//
// Talks to R2 directly over its S3-compatible API using signed requests
// (via aws4fetch), rather than a Cloudflare "R2 binding". This sidesteps
// two separate problems: a dashboard bug where R2 bindings don't save
// correctly under Settings > Bindings, and a shifting wrangler.toml config
// format that's been inconsistent between Cloudflare's own docs and its
// live build output. Plain environment variables are a much older, more
// stable part of the dashboard.
//
// Requires FIVE environment variables set in the Cloudflare Pages
// dashboard under Settings > Environment variables (NOT Bindings):
//   R2_ACCOUNT_ID      -> your Cloudflare account ID
//   R2_ACCESS_KEY_ID   -> from an R2 API token (R2 > Manage API Tokens)
//   R2_SECRET_ACCESS_KEY -> from that same API token
//   R2_BUCKET_NAME     -> the exact name of your R2 bucket
//   PUBLIC_R2_URL      -> the bucket's public pub-xxxxxxxx.r2.dev URL

import { AwsClient } from "aws4fetch";

export async function onRequestPost({ request, env }) {
  try {
    const filename = request.headers.get("X-Filename") || `submission-${Date.now()}.xlsx`;
    const key = `${Date.now()}-${filename}`;

    const client = new AwsClient({
      accessKeyId: env.R2_ACCESS_KEY_ID,
      secretAccessKey: env.R2_SECRET_ACCESS_KEY,
      service: "s3",
      region: "auto",
    });

    const uploadUrl = `https://${env.R2_ACCOUNT_ID}.r2.cloudflarestorage.com/${env.R2_BUCKET_NAME}/${encodeURIComponent(key)}`;

    const fileBytes = await request.arrayBuffer();

    const putRes = await client.fetch(uploadUrl, {
      method: "PUT",
      body: fileBytes,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
    });

    if (!putRes.ok) {
      const detail = await putRes.text();
      return new Response(JSON.stringify({ error: `R2 upload failed: ${putRes.status} ${detail}` }), {
        status: 502,
        headers: { "Content-Type": "application/json" },
      });
    }

    const publicBase = (env.PUBLIC_R2_URL || "").replace(/\/$/, "");
    const url = `${publicBase}/${encodeURIComponent(key)}`;

    return new Response(JSON.stringify({ url }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: String(err) }), {
      status: 500,
      headers: { "Content-Type": "application/json" },
    });
  }
}
