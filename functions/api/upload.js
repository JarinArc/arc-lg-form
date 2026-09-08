// Cloudflare Pages Function — lives at /api/upload once deployed.
// Receives the raw .xlsx bytes from the form (same-origin, so no CORS
// concerns at all), stores them in the bound R2 bucket, and returns a
// public URL. That URL — not the file itself — is what gets sent to the
// Zapier webhook, so there's no base64/binary handling for Zapier to get
// wrong: it just fetches the file from a normal link, the way any file
// download works.
//
// Requires two things configured in the Cloudflare Pages dashboard
// (Settings > Functions):
//   1. An R2 bucket binding named SUBMISSIONS_BUCKET
//   2. An environment variable PUBLIC_R2_URL set to the bucket's public
//      URL (the "Public R2.dev bucket URL" shown after enabling public
//      access on the bucket, or your own custom domain if you set one up)

export async function onRequestPost({ request, env }) {
  try {
    const filename = request.headers.get("X-Filename") || `submission-${Date.now()}.xlsx`;
    const key = `${Date.now()}-${filename}`;

    await env.SUBMISSIONS_BUCKET.put(key, request.body, {
      httpMetadata: {
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
    });

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
