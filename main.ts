// Cortana token server — runs on Deno Deploy (free tier).
// Replaces server.py so your phone can connect without any computer running.
// Your agent itself runs separately on LiveKit Cloud.
//
// Required environment variables (set these in Deno Deploy's dashboard under
// your app's Settings > Environment Variables, NOT written in this file):
//   LIVEKIT_URL          e.g. wss://your-project.livekit.cloud
//   LIVEKIT_API_KEY
//   LIVEKIT_API_SECRET
//   ACCESS_CODE           a password you choose, so strangers can't use your agent

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

function b64url(bytes: Uint8Array): string {
  let str = "";
  for (const b of bytes) str += String.fromCharCode(b);
  return btoa(str).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

async function makeLiveKitToken(
  { apiKey, apiSecret, identity, room }:
  { apiKey: string; apiSecret: string; identity: string; room: string },
): Promise<string> {
  const header = { alg: "HS256", typ: "JWT" };
  const now = Math.floor(Date.now() / 1000);
  const payload = {
    iss: apiKey,
    sub: identity,
    name: "You",
    iat: now,
    nbf: now,
    exp: now + 60 * 60, // 1 hour
    video: { roomJoin: true, room },
    roomConfig: { agents: [{ agentName: "cortana" }] },
  };

  const enc = (obj: unknown) => b64url(new TextEncoder().encode(JSON.stringify(obj)));
  const unsigned = `${enc(header)}.${enc(payload)}`;

  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(apiSecret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"],
  );
  const sigBuf = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(unsigned));
  const sig = b64url(new Uint8Array(sigBuf));

  return `${unsigned}.${sig}`;
}

Deno.serve(async (req: Request) => {
  const url = new URL(req.url);

  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  if (url.pathname !== "/token") {
    return new Response("Not found", { status: 404, headers: corsHeaders });
  }

  const expected = Deno.env.get("ACCESS_CODE") || "";
  const given = url.searchParams.get("code") || "";
  if (!expected || given !== expected) {
    return new Response(JSON.stringify({ error: "wrong or missing access code" }), {
      status: 401,
      headers: { "Content-Type": "application/json", ...corsHeaders },
    });
  }

  const room = "cortana-" + crypto.randomUUID().slice(0, 8);
  const identity = "user-" + crypto.randomUUID().slice(0, 6);

  const token = await makeLiveKitToken({
    apiKey: Deno.env.get("LIVEKIT_API_KEY") || "",
    apiSecret: Deno.env.get("LIVEKIT_API_SECRET") || "",
    identity,
    room,
  });

  return new Response(
    JSON.stringify({ url: Deno.env.get("LIVEKIT_URL") || "", token }),
    { headers: { "Content-Type": "application/json", ...corsHeaders } },
  );
});
