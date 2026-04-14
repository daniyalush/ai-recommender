export const config = {
  runtime: "nodejs"
};

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const body = req.body || {};

    const prompt = `
You are an expert student counselor.

Analyze the student profile and recommend EXACTLY the best 3 study options.

Return ONLY JSON in this format:

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

    const response = await fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=" + process.env.GEMINI_API_KEY,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          contents: [
            {
              parts: [{ text: prompt }]
            }
          ]
        })
      }
    );

    const data = await response.json();

    const text = data.candidates?.[0]?.content?.parts?.[0]?.text || "{}";

    let parsed;

    try {
      parsed = JSON.parse(text);
    } catch (e) {
      parsed = { recommendations: [] };
    }

    res.status(200).json(parsed);

  } catch (error) {
    console.error(error);
    res.status(500).json({
      error: "Failed to generate recommendations"
    });
  }
}