function setCors(res, origin = "*") {
  res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Access-Control-Max-Age", "86400");
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function renderStatusPage({ endpoint, status, details = "" }) {
  const isRunning = status === "running";

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>AI Recommender API Status</title>
  <style>
    body {
      margin: 0;
      padding: 40px 20px;
      font-family: Inter, Arial, sans-serif;
      background: #f7fbff;
      color: #0b1f33;
    }

    .wrap {
      max-width: 760px;
      margin: 0 auto;
    }

    .card {
      position: relative;
      border-radius: 24px;
      padding: 32px;
      background: rgba(255,255,255,0.96);
      box-shadow: 0 20px 50px rgba(11,31,51,0.08);
    }

    .card::before {
      content: "";
      position: absolute;
      inset: 0;
      padding: 2px;
      border-radius: 24px;
      background: linear-gradient(90deg, rgba(6,147,227,0.9), rgba(28,142,237,0.7));
      -webkit-mask:
        linear-gradient(#fff 0 0) content-box,
        linear-gradient(#fff 0 0);
      -webkit-mask-composite: xor;
      mask-composite: exclude;
      pointer-events: none;
    }

    h1 {
      margin: 0 0 12px;
      font-size: 32px;
      line-height: 1.2;
    }

    p {
      margin: 0 0 16px;
      color: #555;
      line-height: 1.6;
      font-size: 15px;
    }

    .status {
      display: inline-block;
      padding: 10px 16px;
      border-radius: 999px;
      font-weight: 700;
      font-size: 14px;
      margin-bottom: 20px;
      background: ${isRunning ? "rgba(16,185,129,0.12)" : "rgba(239,68,68,0.12)"};
      color: ${isRunning ? "#047857" : "#b91c1c"};
    }

    .endpoint-box {
      background: #f3f8fd;
      border: 1px solid rgba(28,142,237,0.18);
      border-radius: 14px;
      padding: 14px 16px;
      word-break: break-all;
      font-size: 14px;
      color: #0b1f33;
      margin: 18px 0 12px;
    }

    .actions {
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      margin-top: 18px;
    }

    button,
    a {
      display: inline-block;
      font-weight: 600;
      padding: 12px 20px;
      font-size: 15px;
      color: #fff;
      background: linear-gradient(90deg, #01b0ef, #1c8eed);
      border-radius: 14px;
      text-decoration: none;
      border: 0;
      cursor: pointer;
      transition: transform 0.25s ease, box-shadow 0.25s ease;
    }

    button:hover,
    a:hover {
      transform: translateY(-2px);
      box-shadow: 0 12px 30px rgba(28,142,237,0.35);
    }

    .note {
      margin-top: 20px;
      font-size: 14px;
      color: #666;
    }

    .details {
      margin-top: 18px;
      padding: 14px 16px;
      border-radius: 14px;
      background: #fff7f7;
      border: 1px solid rgba(239,68,68,0.2);
      color: #991b1b;
      font-size: 14px;
      white-space: pre-wrap;
      word-break: break-word;
    }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <div class="status">${isRunning ? "API Running" : "API Check Failed"}</div>
      <h1>AI Recommender Endpoint</h1>
      <p>This page confirms whether your API endpoint is live and ready to receive Shopify requests.</p>

      <div class="endpoint-box" id="endpointText">${escapeHtml(endpoint)}</div>

      <div class="actions">
        <button onclick="copyEndpoint()">Copy this link</button>
        <a href="${escapeHtml(endpoint)}" target="_blank" rel="noopener">Open endpoint</a>
      </div>

      <p class="note">
        Shopify should use this exact endpoint URL in your section settings.
      </p>

      ${
        details
          ? `<div class="details">${escapeHtml(details)}</div>`
          : ""
      }
    </div>
  </div>

  <script>
    async function copyEndpoint() {
      const text = document.getElementById("endpointText").innerText;
      try {
        await navigator.clipboard.writeText(text);
        alert("Endpoint copied: " + text);
      } catch (err) {
        alert("Copy failed. Please copy manually:\\n" + text);
      }
    }
  </script>
</body>
</html>`;
}

async function checkGeminiApiKey() {
  try {
    return Boolean(process.env.GEMINI_API_KEY);
  } catch {
    return false;
  }
}

export default async function handler(req, res) {
  const allowedOrigins = [
    "https://tristar-education.myshopify.com"
    // Add your custom domain here later:
    // "https://www.tristareducation.com"
  ];

  const requestOrigin = req.headers.origin;
  const origin = allowedOrigins.includes(requestOrigin)
    ? requestOrigin
    : allowedOrigins[0];

  setCors(res, origin);

  const host = req.headers["x-forwarded-host"] || req.headers.host || "";
  const protocol = req.headers["x-forwarded-proto"] || "https";
  const endpoint = `${protocol}://${host}/api/recommend`;

  if (req.method === "OPTIONS") {
    return res.status(204).end();
  }

  if (req.method === "GET") {
    const hasKey = await checkGeminiApiKey();

    return res
      .status(hasKey ? 200 : 500)
      .setHeader("Content-Type", "text/html; charset=utf-8")
      .send(
        renderStatusPage({
          endpoint,
          status: hasKey ? "running" : "failed",
          details: hasKey
            ? "Gemini API key detected. POST requests should work."
            : "GEMINI_API_KEY is missing in your Vercel environment variables."
        })
      );
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const body = req.body || {};

    const prompt = `
You are an expert student counselor.

Analyze the student profile and recommend EXACTLY the best 3 study options.

Return ONLY valid JSON in this format:
{
  "recommendations": [
    {
      "title": "",
      "institution": "",
      "country": "",
      "tuition": "",
      "match_score": "",
      "tags": [],
      "reasons": []
    }
  ]
}

Student:
${JSON.stringify(body)}

Available programs:
[
  {
    "title": "BSc Computer Science",
    "institution": "University of Hertfordshire",
    "country": "UK",
    "tuition": "£15,000/year"
  },
  {
    "title": "BSc Software Engineering",
    "institution": "University of Greenwich",
    "country": "UK",
    "tuition": "£14,500/year"
  },
  {
    "title": "BSc Information Technology",
    "institution": "Middlesex University",
    "country": "UK",
    "tuition": "£14,000/year"
  },
  {
    "title": "MSc Business Analytics",
    "institution": "De Montfort University",
    "country": "UK",
    "tuition": "£16,500/year"
  }
]
`;

    const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
  method: "POST",
  headers: {
    "Authorization": "Bearer " + process.env.GROQ_API_KEY,
    "Content-Type": "application/json"
  },
  body: JSON.stringify({
    model: "llama-3.3-70b-versatile",
    messages: [
      {
        role: "system",
        content: "You are an expert student counselor. Return JSON only."
      },
      {
        role: "user",
        content: prompt
      }
    ]
  })
});

const data = await response.json();

// 🔴 IMPORTANT DEBUG LINE
console.log("GROQ RAW RESPONSE:", data);

// ✅ SAFE PARSE
if (!data.choices || !data.choices.length) {
  return res.status(500).json({
    error: "Groq response invalid",
    raw: data
  });
}

const text = data.choices[0].message.content;

let parsed;
try {
  parsed = JSON.parse(text);
} catch {
  return res.status(500).json({
    error: "AI did not return valid JSON",
    raw_text: text
  });
}

return res.status(200).json(parsed);
  } catch (error) {
    return res.status(500).json({
      error: error.message || "Server error"
    });
  }
}