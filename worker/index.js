// This is the single Worker entry point for Cloudflare's newer unified
// "Workers with static assets" model — replacing the classic Cloudflare
// Pages Functions auto-detection (a `functions/` folder), which is
// specifically a Pages-product feature and isn't picked up the same way
// once a project is created through the newer unified flow.
//
// Every request hits this script first. If the path is /api/upload, we
// handle it ourselves (same R2 upload logic as before). Otherwise, we
// fall through to the ASSETS binding, which serves the built Vite site
// from ./dist exactly as static hosting would.
//
// Requires FIVE environment variables set in the Cloudflare dashboard
// under this Worker's Settings > Variables and Secrets:
//   R2_ACCOUNT_ID        -> your Cloudflare account ID
//   R2_ACCESS_KEY_ID     -> from an R2 API token (R2 > Manage API Tokens)
//   R2_SECRET_ACCESS_KEY -> from that same API token
//   R2_BUCKET_NAME       -> the exact name of your R2 bucket
//   PUBLIC_R2_URL        -> the bucket's public pub-xxxxxxxx.r2.dev URL

import { AwsClient } from "aws4fetch";

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    if (url.pathname === "/api/upload" && request.method === "POST") {
      return handleUpload(request, env);
    }

    // Everything else: serve the built static site.
    return env.ASSETS.fetch(request);
  },
};

async function handleUpload(request, env) {
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
    const publicUrl = `${publicBase}/${encodeURIComponent(key)}`;

    return new Response(JSON.stringify({ url: publicUrl }), {
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
