import fs from "fs";
import path from "path";
import Busboy from "busboy";
import pdfParse from "pdf-parse";
import mammoth from "mammoth";

export const config = {
  api: {
    bodyParser: false
  }
};

function setCors(res, origin = "*") {
  res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Access-Control-Max-Age", "86400");
}

function normalizeText(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim();
}

function safeLower(value) {
  return normalizeText(value).toLowerCase();
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const cleaned = String(value).replace(/[^0-9.]/g, "");
  const num = parseFloat(cleaned);
  return Number.isFinite(num) ? num : null;
}

function parseMultipart(req) {
  return new Promise((resolve, reject) => {
    const bb = Busboy({ headers: req.headers });
    const fields = {};
    const files = [];

    bb.on("field", (name, val) => {
      fields[name] = val;
    });

    bb.on("file", (name, file, info) => {
      const chunks = [];
      file.on("data", (data) => chunks.push(data));
      file.on("end", () => {
        files.push({
          fieldname: name,
          filename: info.filename,
          mimeType: info.mimeType,
          buffer: Buffer.concat(chunks)
        });
      });
    });

    bb.on("finish", () => resolve({ fields, files }));
    bb.on("error", reject);

    req.pipe(bb);
  });
}

async function extractTextFromFile(file) {
  const name = safeLower(file.filename);

  try {
    if (name.endsWith(".pdf")) {
      const parsed = await pdfParse(file.buffer);
      return parsed.text || "";
    }

    if (name.endsWith(".docx")) {
      const parsed = await mammoth.extractRawText({ buffer: file.buffer });
      return parsed.value || "";
    }

    if (name.endsWith(".txt")) {
      return file.buffer.toString("utf8");
    }

    return "";
  } catch {
    return "";
  }
}

function inferFromDocs(docText) {
  const text = safeLower(docText);

  const inferred = {
    education_system: "",
    intended_program: "",
    grades_profile: "",
    english_test: "",
    english_score: "",
    sat_score: ""
  };

  if (text.includes("a level") || text.includes("as level")) inferred.education_system = "A Levels";
  else if (text.includes("o level")) inferred.education_system = "O/A Levels";
  else if (text.includes("fsc")) inferred.education_system = "FSC";
  else if (text.includes("ib")) inferred.education_system = "IB";

  const majorPatterns = [
    ["business administration", "Business Administration"],
    ["bba", "Business Administration"],
    ["business", "Business Administration"],
    ["management", "Business Administration"],
    ["economics", "Economics"],
    ["finance", "Economics and Finance"],
    ["international trade", "International Trade and Business"],
    ["computer science", "Computer Engineering"],
    ["computer engineering", "Computer Engineering"],
    ["software engineering", "Software Engineering"],
    ["artificial intelligence", "Artificial Intelligence Engineering"],
    ["psychology", "Psychology"],
    ["medicine", "Medicine"],
    ["dentistry", "Dentistry"],
    ["nursing", "Nursing"],
    ["architecture", "Architecture"]
  ];

  for (const [pattern, label] of majorPatterns) {
    if (text.includes(pattern)) {
      inferred.intended_program = label;
      break;
    }
  }

  const ieltsMatch = docText.match(/IELTS[^0-9]*([0-9](?:\.[0-9])?)/i);
  if (ieltsMatch) {
    inferred.english_test = "IELTS";
    inferred.english_score = ieltsMatch[1];
  }

  const satMatch = docText.match(/SAT[^0-9]*([0-9]{3,4})/i);
  if (satMatch) {
    inferred.sat_score = satMatch[1];
  }

  inferred.grades_profile = normalizeText(docText).slice(0, 800);

  return inferred;
}

function loadPrograms() {
  const filePath = path.join(process.cwd(), "knowledge", "Programs.json");
  if (!fs.existsSync(filePath)) {
    throw new Error("Programs.json not found in knowledge folder");
  }

  const raw = JSON.parse(fs.readFileSync(filePath, "utf8"));

  return raw
    .map((row) => ({
      program_title: normalizeText(row.Program_Title),
      major_tag: normalizeText(row.Major_Tag),
      degree_level: normalizeText(row.Degree_Level),
      speciality: normalizeText(row.Speciality),
      university: normalizeText(row.University),
      city: normalizeText(row.City),
      campus_style: normalizeText(row.Campus_Style),
      intake: normalizeText(row.Intake),
      scholarship_tag: normalizeText(row.Scholarship_Tag),
      application_fee_tag: normalizeText(row.Application_Fee_Tag),
      tuition_usd: toNumber(row.Tuition_USD),

      tuition_basis: normalizeText(row.Tuition_Basis),
      duration_years_assumed: toNumber(row.Duration_Years_Assumed),
      category_tier: normalizeText(row.Category_Tier),
      ranking_band: normalizeText(row.Ranking_Band),

      raw_title: normalizeText(row.Source_Raw_Title || row.Raw_Title),
      raw_tags: normalizeText(row.Source_Raw_Tags || row.Raw_Tags)
    }))
    .filter((row) => {
      const level = safeLower(row.degree_level);
      return row.university && row.program_title && level.includes("bachelor");
    });
}

const knowledge = {
  rules: {
    country_scope: "Turkey only",
    intake: "Fall only",
    sat_policy: "SAT is not mandatory for most Turkish universities but may help scholarships",
    deposit: "Seat reservation deposit is usually around $1,000",
    living_costs: {
      dorm_total_year: "$4,400 to $8,000",
      shared_apartment_total_year: "$6,400 to $10,000",
      food_transport_month: "$300 to $500"
    }
  },
  categories: {
    top_tier: ["Sabanci University", "Koc University"],
    strong_mid_tier: [
      "Ozyegin University",
      "Kadir Has University",
      "Istanbul Bilgi University",
      "Bahcesehir University (BAU)",
      "Bahcesehir University",
      "Istinye University"
    ],
    budget_friendly: [
      "Istanbul Medipol University",
      "Isik University",
      "TED University",
      "Yasar University",
      "Uskudar University",
      "Istanbul Aydin University"
    ],
    non_preferred: [
      "Istanbul Okan University",
      "Altinbas University"
    ],
    extreme_budget: [
      "Beykoz University",
      "Kent University",
      "Istanbul Kent University",
      "beykent University",
      "Beykent University",
      "Atlas University"
    ]
  },
  ranking_overrides: {
    "Sabanci University": "THE 351–400 band",
    "Koc University": "THE 301–350 band",
    "Kadir Has University": "THE 601–800 band",
    "Ozyegin University": "THE 801–1000 band",
    "Bahcesehir University (BAU)": "THE 801–1000 band",
    "Bahcesehir University": "THE 801–1000 band",
    "Istanbul Medipol University": "THE 801–1000 band",
    "Istanbul Bilgi University": "THE 1200+ band",
    "TED University": "Limited ranking presence"
  },
  programs: loadPrograms()
};

function categorizeUniversity(programRow) {
  if (programRow.category_tier) return programRow.category_tier;

  const n = safeLower(programRow.university);

  for (const item of knowledge.categories.top_tier) {
    if (safeLower(item) === n) return "Top Tier";
  }
  for (const item of knowledge.categories.strong_mid_tier) {
    if (safeLower(item) === n) return "Strong Mid-Tier";
  }
  for (const item of knowledge.categories.budget_friendly) {
    if (safeLower(item) === n) return "Budget-Friendly";
  }
  for (const item of knowledge.categories.non_preferred) {
    if (safeLower(item) === n) return "Non-Preferred";
  }
  for (const item of knowledge.categories.extreme_budget) {
    if (safeLower(item) === n) return "Extreme Budget";
  }

  return "General";
}

function categoryBaseScore(category) {
  const scores = {
    "Top Tier": 40,
    "Strong Mid-Tier": 30,
    "Budget-Friendly": 20,
    "General": 10,
    "Extreme Budget": 5,
    "Non-Preferred": -10
  };
  return scores[category] ?? 0;
}

function parseProgramIntent(input) {
  const q = safeLower(input);

  const degree = {
    wants_ba: /\bba\b|\bbachelor of arts\b/.test(q),
    wants_bs: /\bbs\b|\bbsc\b|\bbachelor of science\b/.test(q),
    wants_bba: /\bbba\b|\bbachelor of business administration\b/.test(q)
  };

  let family = [];

  if (
    q.includes("computer science") ||
    q.includes("cs") ||
    q.includes("computer engineering")
  ) {
    family = [
      "computer engineering",
      "computer science",
      "software engineering"
    ];
  } else if (
    q.includes("electrical") ||
    q.includes("electronics engineering") ||
    q.includes("electrical & electronics")
  ) {
    family = [
      "electrical & electronics engineering",
      "electrical - electronics engineering",
      "electrical electronics engineering"
    ];
  } else if (
    q.includes("business administration") ||
    q.includes("bba") ||
    q === "business" ||
    q.includes("management")
  ) {
    family = [
      "business administration"
    ];
  } else if (
    q.includes("economics") ||
    q.includes("finance")
  ) {
    family = [
      "economics",
      "economics & finance",
      "economics and finance",
      "finance & banking",
      "international trade & finance"
    ];
  } else if (q.includes("psychology")) {
    family = ["psychology"];
  } else if (q.includes("software")) {
    family = ["software engineering", "software development"];
  } else if (q.includes("data science")) {
    family = ["data science", "data engineering", "analytics"];
  }

  return {
    raw: q,
    family,
    ...degree
  };
}

function titleDegreeMatch(programTitle, intent) {
  const title = safeLower(programTitle);

  if (intent.wants_bba) {
    return title.includes("business administration");
  }

  if (intent.wants_bs) {
    return (
      title.includes("(bsc)") ||
      title.includes("bachelor of science") ||
      title.includes("bachelors of science")
    );
  }

  if (intent.wants_ba) {
    return (
      title.includes("(ba)") ||
      title.includes("bachelor of arts") ||
      title.includes("bachelors of arts")
    );
  }

  return true;
}

function strictProgramFilter(row, intendedProgram) {
  const intent = parseProgramIntent(intendedProgram);

  const title = safeLower(row.program_title);
  const major = safeLower(row.major_tag);
  const combined = `${title} ${major}`;

  if (!titleDegreeMatch(row.program_title, intent)) {
    return false;
  }

  if (intent.family.length) {
    const positive = intent.family.some(term => combined.includes(term));
    if (!positive) return false;
  }

  const hardNegatives = [
    "first and emergency aid",
    "comparative literature",
    "political science",
    "aviation management",
    "anthropology",
    "theatre",
    "translation",
    "gastronomy",
    "history",
    "philosophy",
    "archaeology"
  ];

  if (intent.raw.includes("electrical")) {
    if (!combined.includes("electrical")) return false;
  }

  if (
    intent.raw.includes("cs") ||
    intent.raw.includes("computer science") ||
    intent.raw.includes("computer engineering")
  ) {
    if (
      combined.includes("data science") ||
      combined.includes("analytics")
    ) {
      return false;
    }
  }

  if (hardNegatives.some(term => combined.includes(term))) {
    if (
      intent.raw.includes("business") ||
      intent.raw.includes("bba") ||
      intent.raw.includes("cs") ||
      intent.raw.includes("electrical")
    ) {
      return false;
    }
  }

  return true;
}

function programMatchScore(programRow, intendedProgram) {
  const intent = parseProgramIntent(intendedProgram);
  const title = safeLower(programRow.program_title);
  const major = safeLower(programRow.major_tag);
  const speciality = safeLower(programRow.speciality);
  const haystack = `${title} ${major} ${speciality}`;

  let score = 0;

  for (const target of intent.family) {
    if (major.includes(target)) score += 30;
    if (title.includes(target)) score += 25;
    if (speciality.includes(target)) score += 10;
  }

  if (intent.wants_bba && title.includes("business administration")) score += 20;
  if (intent.wants_bs && (title.includes("(bsc)") || title.includes("bachelor of science") || title.includes("bachelors of science"))) score += 20;
  if (intent.wants_ba && (title.includes("(ba)") || title.includes("bachelor of arts") || title.includes("bachelors of arts"))) score += 20;

  if (title.includes("phd")) score -= 1000;
  if (title.includes("master")) score -= 1000;
  if (title.includes("doctoral")) score -= 1000;
  if (title.includes("associate")) score -= 100;

  return score;
}

function rankingScore(university, rankingImportance) {
  if (safeLower(rankingImportance) !== "yes") return 0;
  const category = categorizeUniversity(university);
  if (category === "Top Tier") return 12;
  if (category === "Strong Mid-Tier") return 8;
  if (category === "Budget-Friendly") return 4;
  return 0;
}

function scholarshipScore(programRow, scholarshipPreference) {
  if (safeLower(scholarshipPreference) !== "yes") return 0;
  return safeLower(programRow.scholarship_tag).includes("yes") ? 6 : 0;
}

function lifestyleScore(university, lifestylePreference) {
  const pref = safeLower(lifestylePreference);
  const uni = safeLower(university);

  const campusUnis = ["sabanci university", "ozyegin university", "isik university", "koc university"];
  const cityUnis = [
    "istanbul bilgi university",
    "kadir has university",
    "bahcesehir university",
    "bahcesehir university (bau)",
    "istanbul medipol university",
    "istanbul aydin university",
    "istanbul okan university",
    "atlas university",
    "beykent university",
    "beykent university",
    "istanbul kent university",
    "uskudar university",
    "istinye university"
  ];

  if (pref.includes("campus")) {
    return campusUnis.some((x) => uni.includes(x)) ? 8 : 0;
  }

  if (pref.includes("city")) {
    return cityUnis.some((x) => uni.includes(x)) ? 8 : 0;
  }

  return 0;
}

function budgetScore(programRow, yearlyBudgetTotal, lifestylePreference) {
  const budget = toNumber(yearlyBudgetTotal);
  if (!budget || !programRow.tuition_usd) return 0;

  const livingLow = safeLower(lifestylePreference).includes("campus") ? 4400 : 6400;
  const estimatedTotal = programRow.tuition_usd + livingLow;

  if (estimatedTotal <= budget * 0.8) return 10;
  if (estimatedTotal <= budget) return 6;
  if (estimatedTotal <= budget * 1.15) return 2;
  return -8;
}

function preferenceUniversityScore(university, universityPreference) {
  const pref = safeLower(universityPreference);
  if (!pref) return 0;
  return safeLower(university).includes(pref) ? 12 : 0;
}

function docsOnlyMissingFields(merged) {
  return [
    { key: "ranking_importance", label: "Is university ranking an important factor for you?" },
    { key: "yearly_budget_total", label: "What is your approximate yearly budget for tuition and living?" },
    { key: "scholarship_preference", label: "Are you specifically looking for high scholarships?" },
    { key: "lifestyle_preference", label: "Do you prefer a city-center university experience or a campus-based environment?" },
    { key: "candidate_type", label: "Are you a private candidate or applying through a school?" }
  ].filter((f) => !String(merged[f.key] || "").trim());
}

function fullMissingFields(merged) {
  return [
    { key: "university_preference", label: "Do you have any specific university preference?" },
    { key: "ranking_importance", label: "Is university ranking an important factor for you?" },
    { key: "yearly_budget_total", label: "What is your approximate yearly budget for tuition and living?" },
    { key: "scholarship_preference", label: "Are you specifically looking for high scholarships?" },
    { key: "lifestyle_preference", label: "Do you prefer a city-center university experience or a campus-based environment?" },
    { key: "education_system", label: "What education system are you applying with?" },
    { key: "sat_score", label: "Do you have an SAT score? If yes, what score?" },
    { key: "intended_program", label: "What is your intended program/major?" },
    { key: "grades_profile", label: "Are you in final year or already completed? Share grades/predicted." },
    { key: "candidate_type", label: "Are you a private candidate or applying through a school?" }
  ].filter((f) => !String(merged[f.key] || "").trim());
}

function buildQuestionResponse(missing, docsUploaded) {
  return {
    mode: "questions",
    country_scope: "Turkey only",
    docs_uploaded: docsUploaded,
    message: "Please answer these before I recommend universities.",
    questions: missing.map((f) => ({
      key: f.key,
      question: f.label
    }))
  };
}

function scorePrograms(student) {
  const filtered = knowledge.programs.filter(function(row) {
    return strictProgramFilter(row, student.intended_program || "");
  });

  const pool = filtered.length ? filtered : knowledge.programs;

  const scored = pool.map((row) => {
    const category = categorizeUniversity(row);

    let score = 0;
    score += categoryBaseScore(category);
    score += programMatchScore(row, student.intended_program);
    score += preferenceUniversityScore(row.university, student.university_preference);
    score += rankingScore(row.university, student.ranking_importance);
    score += scholarshipScore(row, student.scholarship_preference);
    score += lifestyleScore(row.university, student.lifestyle_preference);
    score += budgetScore(row, student.yearly_budget_total, student.lifestyle_preference);

    return {
      ...row,
      _category: category,
      _score: score
    };
  });

  scored.sort((a, b) => b._score - a._score);

  const seenUniversities = new Set();
  const finalRows = [];

  for (const row of scored) {
    if (finalRows.length >= 3) break;
    if (row._score < 15) continue;

    const uniKey = safeLower(row.university);
    if (seenUniversities.has(uniKey)) continue;

    seenUniversities.add(uniKey);
    finalRows.push(row);
  }

  return finalRows;
}

function livingEstimate(lifestylePreference) {
  const pref = safeLower(lifestylePreference);
  if (pref.includes("campus")) return knowledge.rules.living_costs.dorm_total_year;
  if (pref.includes("city")) return knowledge.rules.living_costs.shared_apartment_total_year;
  return "$4,400 to $10,000";
}

function buildRecommendationResponse(bestRows, student, docsUsed) {
  return {
    mode: "recommendations",
    country_scope: "Turkey only",
    docs_used: docsUsed,
    summary: {
      budget_comment: `Recommendations were matched against an approximate yearly tuition + living budget of ${student.yearly_budget_total || "not fully specified"}.`,
      ranking_comment: safeLower(student.ranking_importance) === "yes"
        ? "Ranking-sensitive options were prioritized."
        : "Ranking was not treated as the primary driver.",
      scholarship_comment: "Scholarships are shown as possible/estimated, not guaranteed.",
      lifestyle_comment: student.lifestyle_preference
        ? `Lifestyle preference considered: ${student.lifestyle_preference}.`
        : "Lifestyle preference was not fully specified.",
      intake_comment: "For undergraduate Türkiye admissions, the intake is Fall only."
    },
    recommendations: bestRows.map((row) => ({
      university: row.university,
      program: row.program_title,
      category: row._category,
      ranking_band: row.ranking_band || knowledge.ranking_overrides[row.university] || "Indicative only",
      tuition_estimate: row.tuition_usd ? `$${row.tuition_usd}/year` : "Ask counselor",
      living_cost_estimate: livingEstimate(student.lifestyle_preference),
      estimated_total_yearly_cost: row.tuition_usd
        ? `Approx. $${row.tuition_usd} tuition + living depending on lifestyle`
        : "Depends on tuition + living style",
      scholarship_positioning: safeLower(row.scholarship_tag).includes("yes")
        ? "Scholarship/discount may be possible depending on profile and university policy."
        : "Scholarship may still be possible depending on profile and university policy.",
      why_this_fits: [
        `Strong match for the intended program: ${student.intended_program}.`,
        `Selected within the ${row._category} category based on TriStar counseling positioning.`,
        `Matched against budget, lifestyle, ranking, and scholarship preference.`
      ],
      notes: [
  "Undergraduate Türkiye intake is Fall only.",
  row.tuition_basis === "Annualized Estimate"
    ? "Tuition shown here is an annualized estimate converted from a whole-program fee."
    : "Tuition shown here is treated as an annual fee.",
  "Seat reservation deposits are commonly around $1,000, depending on the university."
]
    }))
  };
}

export default async function handler(req, res) {
  const allowedOrigins = [
    "https://tristar-education.myshopify.com",
    "https://tsapply.online"
  ];

  const requestOrigin = req.headers.origin;
  const origin = allowedOrigins.includes(requestOrigin) ? requestOrigin : allowedOrigins[0];
  setCors(res, origin);

  if (req.method === "OPTIONS") {
    return res.status(204).end();
  }

  if (req.method === "GET") {
    return res.status(200).json({
      ok: true,
      country_scope: "Turkey only",
      loaded_program_rows: knowledge.programs.length
    });
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    let body = {};
    let files = [];

    const contentType = req.headers["content-type"] || "";

    if (contentType.includes("multipart/form-data")) {
      const parsed = await parseMultipart(req);
      body = parsed.fields || {};
      files = parsed.files || [];
    } else {
      await new Promise((resolve, reject) => {
        let raw = "";
        req.on("data", (chunk) => { raw += chunk; });
        req.on("end", () => {
          try {
            body = raw ? JSON.parse(raw) : {};
            resolve();
          } catch (e) {
            reject(e);
          }
        });
        req.on("error", reject);
      });
    }

    let extractedText = "";
    for (const file of files) {
      const text = await extractTextFromFile(file);
      if (text) {
        extractedText += `\n\n[FILE: ${file.filename}]\n${text}`;
      }
    }

    const inferred = extractedText ? inferFromDocs(extractedText) : {};

    const merged = {
      ...body,
      education_system: body.education_system || inferred.education_system || "",
      intended_program: body.intended_program || body.course_interest || inferred.intended_program || "",
      grades_profile: body.grades_profile || inferred.grades_profile || "",
      sat_score: body.sat_score || inferred.sat_score || "",
      english_test: body.english_test || inferred.english_test || "",
      english_score: body.english_score || inferred.english_score || ""
    };

    const docsOnlyMode = files.length > 0 && Object.values(body).every((v) => !String(v || "").trim());

    const missing = docsOnlyMode ? docsOnlyMissingFields(merged) : fullMissingFields(merged);

    if (missing.length > 0) {
      return res.status(200).json(buildQuestionResponse(missing, files.length > 0));
    }

    const bestRows = scorePrograms(merged);

    if (!bestRows.length) {
      return res.status(200).json({
        mode: "recommendations",
        country_scope: "Turkey only",
        docs_used: files.length > 0,
        summary: {
          budget_comment: "No strong match was found with the current inputs.",
          ranking_comment: "",
          scholarship_comment: "",
          lifestyle_comment: "",
          intake_comment: "For undergraduate Türkiye admissions, the intake is Fall only."
        },
        recommendations: []
      });
    }

    return res.status(200).json(buildRecommendationResponse(bestRows, merged, files.length > 0));
  } catch (error) {
    return res.status(500).json({
      error: error.message || "Server error"
    });
  }
}