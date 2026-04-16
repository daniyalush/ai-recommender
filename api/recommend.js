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

const ALLOWED_ORIGINS = [
  "https://tristar-education.myshopify.com",
  "https://tsapply.online"
];

const CATEGORY_BASE_SCORES = {
  "Top Tier": 40,
  "Strong Mid-Tier": 30,
  "Budget-Friendly": 20,
  "General": 10,
  "Extreme Budget": 5,
  "Non-Preferred": -10
};

const UNIVERSITY_CATEGORY_MAP = {
  "sabanci university": "Top Tier",
  "koc university": "Top Tier",
  "ozyegin university": "Strong Mid-Tier",
  "kadir has university": "Strong Mid-Tier",
  "istanbul bilgi university": "Strong Mid-Tier",
  "bahcesehir university (bau)": "Strong Mid-Tier",
  "bahcesehir university": "Strong Mid-Tier",
  "istinye university": "Strong Mid-Tier",
  "istanbul medipol university": "Budget-Friendly",
  "isik university": "Budget-Friendly",
  "ted university": "Budget-Friendly",
  "yasar university": "Budget-Friendly",
  "uskudar university": "Budget-Friendly",
  "istanbul aydin university": "Budget-Friendly",
  "istanbul okan university": "Non-Preferred",
  "altinbas university": "Non-Preferred",
  "beykoz university": "Extreme Budget",
  "kent university": "Extreme Budget",
  "istanbul kent university": "Extreme Budget",
  "beykent university": "Extreme Budget",
  "atlas university": "Extreme Budget"
};

const RANKING_OVERRIDES = {
  "Sabanci University": "THE 351–400 band",
  "Koc University": "THE 301–350 band",
  "Kadir Has University": "THE 601–800 band",
  "Ozyegin University": "THE 801–1000 band",
  "Bahcesehir University (BAU)": "THE 801–1000 band",
  "Bahcesehir University": "THE 801–1000 band",
  "Istanbul Medipol University": "THE 801–1000 band",
  "Istanbul Bilgi University": "THE 1200+ band",
  "TED University": "Limited ranking presence"
};

const PROGRAM_FAMILIES = [
  {
    key: "business_administration",
    label: "Business Administration",
    aliases: [
      "business administration",
      "bba",
      "business",
      "management",
      "business management"
    ],
    degreeBias: "bba",
    positivePatterns: [
      /business administration/,
      /\bbba\b/
    ],
    negativePatterns: [
      /management engineering/,
      /aviation management/,
      /sports management/,
      /logistics management/,
      /management information systems/
    ]
  },
  {
    key: "economics",
    label: "Economics",
    aliases: [
      "economics",
      "economics and finance",
      "economics & finance",
      "finance",
      "banking"
    ],
    positivePatterns: [
      /economics/,
      /economics\s*(and|&)\s*finance/,
      /finance and banking/,
      /finance\s*&\s*banking/,
      /international trade and finance/,
      /international trade\s*&\s*finance/
    ],
    negativePatterns: [
      /computer/,
      /software/,
      /data science/,
      /artificial intelligence/
    ]
  },
  {
    key: "computer",
    label: "Computer Engineering",
    aliases: [
      "computer science",
      "cs",
      "computer engineering",
      "software engineering",
      "software development"
    ],
    positivePatterns: [
      /computer science/,
      /computer engineering/,
      /software engineering/,
      /software development/
    ],
    negativePatterns: [
      /economics/,
      /finance/,
      /psychology/,
      /history/,
      /comparative literature/
    ]
  },
  {
    key: "electrical",
    label: "Electrical & Electronics Engineering",
    aliases: [
      "electrical",
      "electrical engineering",
      "electronics engineering",
      "electrical & electronics",
      "electrical-electronics"
    ],
    positivePatterns: [
      /electrical/,
      /electronics engineering/,
      /electrical\s*&\s*electronics/,
      /electrical-?electronics/
    ],
    negativePatterns: []
  },
  {
    key: "psychology",
    label: "Psychology",
    aliases: ["psychology"],
    positivePatterns: [/psychology/, /psychological counselling/],
    negativePatterns: []
  },
  {
    key: "medicine",
    label: "Medicine",
    aliases: ["medicine", "md", "doctor of medicine"],
    positivePatterns: [/medicine/, /doctor of medicine/, /\bmd\b/],
    negativePatterns: [/dentistry/, /nursing/, /pharmacy/]
  },
  {
    key: "dentistry",
    label: "Dentistry",
    aliases: ["dentistry", "dental", "bds"],
    positivePatterns: [/dentistry/, /dental/, /\bbds\b/],
    negativePatterns: [/medicine/, /nursing/]
  },
  {
    key: "nursing",
    label: "Nursing",
    aliases: ["nursing"],
    positivePatterns: [/nursing/],
    negativePatterns: []
  },
  {
    key: "architecture",
    label: "Architecture",
    aliases: ["architecture", "interior architecture"],
    positivePatterns: [/architecture/],
    negativePatterns: [/archaeology/]
  },
  {
    key: "artificial_intelligence",
    label: "Artificial Intelligence",
    aliases: ["artificial intelligence", "ai", "machine learning"],
    positivePatterns: [/artificial intelligence/, /ai engineering/, /ai and data/],
    negativePatterns: [/economics/, /psychology/]
  },
  {
    key: "data_science",
    label: "Data Science",
    aliases: ["data science", "data analytics", "analytics", "data engineering"],
    positivePatterns: [/data science/, /analytics/, /data engineering/],
    negativePatterns: [/economics/, /finance/, /computer engineering/]
  }
];

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

function normalizeForMatch(value) {
  return safeLower(value)
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function tokenize(value) {
  return normalizeForMatch(value).split(" ").filter(Boolean);
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

  const matchedFamily = matchProgramFamily(text);
  if (matchedFamily) inferred.intended_program = matchedFamily.label;

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
  const candidates = [
    path.join(process.cwd(), "knowledge", "Programs.json"),
    path.join(process.cwd(), "knowledge", "Programs(2).json"),
    path.join(process.cwd(), "Programs.json"),
    path.join(process.cwd(), "Programs(2).json")
  ];

  const existingPath = candidates.find((candidate) => fs.existsSync(candidate));
  if (!existingPath) {
    throw new Error("Programs dataset not found. Expected one of: knowledge/Programs.json, knowledge/Programs(2).json, Programs.json, Programs(2).json");
  }

  const raw = JSON.parse(fs.readFileSync(existingPath, "utf8"));

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
  programs: loadPrograms()
};

function categorizeUniversity(programRow) {
  if (programRow.category_tier) return programRow.category_tier;
  return UNIVERSITY_CATEGORY_MAP[safeLower(programRow.university)] || "General";
}

function categoryBaseScore(category) {
  return CATEGORY_BASE_SCORES[category] ?? 0;
}

function detectRequestedDegree(input) {
  const q = normalizeForMatch(input);
  return {
    wants_ba: /\bba\b|\bbachelor of arts\b/.test(q),
    wants_bs: /\bbs\b|\bbsc\b|\bbachelor of science\b/.test(q),
    wants_bba: /\bbba\b|\bbachelor of business administration\b/.test(q)
  };
}

function matchProgramFamily(input) {
  const q = normalizeForMatch(input);
  if (!q) return null;

  const scored = PROGRAM_FAMILIES.map((family) => {
    let score = 0;

    for (const alias of family.aliases) {
      const aliasNorm = normalizeForMatch(alias);
      if (aliasNorm && q.includes(aliasNorm)) score += aliasNorm.length >= 10 ? 8 : 5;
    }

    for (const pattern of family.positivePatterns) {
      if (pattern.test(q)) score += 12;
    }

    for (const pattern of family.negativePatterns) {
      if (pattern.test(q)) score -= 20;
    }

    return { family, score };
  }).sort((a, b) => b.score - a.score);

  if (!scored.length || scored[0].score <= 0) return null;
  return scored[0].family;
}

function parseProgramIntent(input) {
  const raw = normalizeText(input);
  const normalized = normalizeForMatch(input);
  const degree = detectRequestedDegree(input);
  const family = matchProgramFamily(input);

  return {
    raw,
    normalized,
    family,
    familyKey: family?.key || null,
    familyLabel: family?.label || raw,
    ...degree
  };
}

function getProgramSearchBlob(row) {
  return normalizeForMatch([
    row.program_title,
    row.major_tag,
    row.speciality,
    row.raw_title,
    row.raw_tags
  ].join(" "));
}

function titleDegreeMatch(programTitle, intent) {
  const title = normalizeForMatch(programTitle);

  if (intent.wants_bba) {
    return /business administration|\bbba\b/.test(title);
  }

  if (intent.wants_bs) {
    return /(\bbsc\b|bachelor of science|bachelors of science)/.test(title);
  }

  if (intent.wants_ba) {
    return /(\bba\b|bachelor of arts|bachelors of arts)/.test(title);
  }

  return true;
}

function rowMatchesFamily(row, family) {
  if (!family) return true;
  const blob = getProgramSearchBlob(row);

  const positive = family.positivePatterns.some((pattern) => pattern.test(blob));
  if (!positive) return false;

  const negative = family.negativePatterns.some((pattern) => pattern.test(blob));
  if (negative) return false;

  return true;
}

function strictProgramFilter(row, intendedProgram) {
  const intent = parseProgramIntent(intendedProgram);

  if (!intent.normalized) return true;
  if (!titleDegreeMatch(row.program_title, intent)) return false;
  if (!rowMatchesFamily(row, intent.family)) return false;

  return true;
}

function computeTokenOverlapScore(blob, intendedProgram) {
  const queryTokens = tokenize(intendedProgram).filter((token) => token.length >= 3 && !["bachelor", "arts", "science"].includes(token));
  if (!queryTokens.length) return 0;

  let score = 0;
  for (const token of queryTokens) {
    if (blob.includes(token)) score += 3;
  }
  return score;
}

function programMatchScore(programRow, intendedProgram) {
  const intent = parseProgramIntent(intendedProgram);
  const blob = getProgramSearchBlob(programRow);

  if (!intent.normalized) return 0;
  if (intent.family && !rowMatchesFamily(programRow, intent.family)) return -1000;
  if (!titleDegreeMatch(programRow.program_title, intent)) return -120;

  let score = 0;

  if (intent.family) {
    score += 35;

    for (const pattern of intent.family.positivePatterns) {
      if (pattern.test(normalizeForMatch(programRow.major_tag))) score += 30;
      if (pattern.test(normalizeForMatch(programRow.program_title))) score += 25;
      if (pattern.test(normalizeForMatch(programRow.speciality))) score += 10;
      if (pattern.test(normalizeForMatch(programRow.raw_tags))) score += 8;
    }
  }

  score += computeTokenOverlapScore(blob, intendedProgram);

  if (intent.wants_bba && /business administration|\bbba\b/.test(blob)) score += 18;
  if (intent.wants_bs && /(\bbsc\b|bachelor of science|bachelors of science)/.test(blob)) score += 18;
  if (intent.wants_ba && /(\bba\b|bachelor of arts|bachelors of arts)/.test(blob)) score += 18;

  if (/phd|master|doctoral|associate/.test(blob)) score -= 1000;

  return score;
}

function rankingScore(university, rankingImportance) {
  if (safeLower(rankingImportance) !== "yes") return 0;
  const category = typeof university === "string"
    ? (UNIVERSITY_CATEGORY_MAP[safeLower(university)] || "General")
    : categorizeUniversity(university);
  if (category === "Top Tier") return 12;
  if (category === "Strong Mid-Tier") return 8;
  if (category === "Budget-Friendly") return 4;
  return 0;
}

function scholarshipScore(programRow, scholarshipPreference) {
  if (safeLower(scholarshipPreference) !== "yes") return 0;
  return safeLower(programRow.scholarship_tag).includes("yes") ? 6 : 0;
}

function lifestyleScore(programRow, lifestylePreference) {
  const pref = safeLower(lifestylePreference);
  if (!pref) return 0;

  const campusStyle = safeLower(programRow.campus_style);
  const university = safeLower(programRow.university);

  if (pref.includes("campus")) {
    if (campusStyle.includes("campus")) return 8;
    if (["sabanci university", "ozyegin university", "koc university", "isik university"].includes(university)) return 8;
    return 0;
  }

  if (pref.includes("city")) {
    if (campusStyle.includes("city")) return 8;
    if ([
      "istanbul bilgi university",
      "kadir has university",
      "bahcesehir university",
      "bahcesehir university (bau)",
      "istanbul medipol university",
      "istanbul aydin university",
      "istanbul okan university",
      "atlas university",
      "beykent university",
      "istanbul kent university",
      "uskudar university",
      "istinye university"
    ].includes(university)) return 8;
    return 0;
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
  const pref = normalizeForMatch(universityPreference);
  if (!pref) return 0;
  return normalizeForMatch(university).includes(pref) ? 12 : 0;
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

function buildDiagnostics(row, student, breakdown) {
  return {
    intended_program: student.intended_program || "",
    matched_family: parseProgramIntent(student.intended_program || "").familyLabel,
    score_breakdown: breakdown,
    matching_blob: getProgramSearchBlob(row)
  };
}

function scorePrograms(student) {
  const intent = parseProgramIntent(student.intended_program || "");
  const strictPool = knowledge.programs.filter((row) => strictProgramFilter(row, student.intended_program || ""));

  let pool = strictPool;
  let usedRelaxedFallback = false;

  if (!pool.length) {
    if (intent.family) {
      pool = knowledge.programs.filter((row) => rowMatchesFamily(row, intent.family));
      usedRelaxedFallback = true;
    } else {
      pool = knowledge.programs.filter((row) => titleDegreeMatch(row.program_title, intent));
      usedRelaxedFallback = true;
    }
  }

  const scored = pool
    .map((row) => {
      const category = categorizeUniversity(row);
      const breakdown = {
        category: categoryBaseScore(category),
        program_match: programMatchScore(row, student.intended_program),
        university_preference: preferenceUniversityScore(row.university, student.university_preference),
        ranking: rankingScore(row.university, student.ranking_importance),
        scholarship: scholarshipScore(row, student.scholarship_preference),
        lifestyle: lifestyleScore(row, student.lifestyle_preference),
        budget: budgetScore(row, student.yearly_budget_total, student.lifestyle_preference)
      };

      const total = Object.values(breakdown).reduce((sum, value) => sum + value, 0);

      return {
        ...row,
        _category: category,
        _score: total,
        _diagnostics: buildDiagnostics(row, student, breakdown)
      };
    })
    .filter((row) => row._score > -500)
    .sort((a, b) => {
      if (b._score !== a._score) return b._score - a._score;
      if ((a.tuition_usd ?? Infinity) !== (b.tuition_usd ?? Infinity)) {
        return (a.tuition_usd ?? Infinity) - (b.tuition_usd ?? Infinity);
      }
      return a.university.localeCompare(b.university);
    });

  const finalRows = [];
  const seenProgramKeys = new Set();

  for (const row of scored) {
    if (finalRows.length >= 5) break;
    if (row._score < 20) continue;

    const key = `${safeLower(row.university)}::${safeLower(row.program_title)}`;
    if (seenProgramKeys.has(key)) continue;

    seenProgramKeys.add(key);
    finalRows.push(row);
  }

  return {
    rows: finalRows,
    diagnostics: {
      strict_matches: strictPool.length,
      pool_size: pool.length,
      used_relaxed_fallback: usedRelaxedFallback,
      resolved_family: intent.familyLabel || "Unresolved"
    }
  };
}

function livingEstimate(lifestylePreference) {
  const pref = safeLower(lifestylePreference);
  if (pref.includes("campus")) return knowledge.rules.living_costs.dorm_total_year;
  if (pref.includes("city")) return knowledge.rules.living_costs.shared_apartment_total_year;
  return "$4,400 to $10,000";
}

function explainWhyThisFits(row, student) {
  const reasons = [];
  const intent = parseProgramIntent(student.intended_program || "");

  if (intent.familyLabel) {
    reasons.push(`Matched to intended program family: ${intent.familyLabel}.`);
  }

  reasons.push(`Selected within the ${row._category} category based on TriStar counseling positioning.`);

  if (student.yearly_budget_total) {
    reasons.push(`Compared against the provided yearly budget of ${student.yearly_budget_total}.`);
  }

  if (student.lifestyle_preference) {
    reasons.push(`Considered lifestyle preference: ${student.lifestyle_preference}.`);
  }

  return reasons;
}

function buildRecommendationResponse(bestRows, student, docsUsed, diagnostics) {
  return {
    mode: "recommendations",
    country_scope: "Turkey only",
    docs_used: docsUsed,
    diagnostics,
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
      ranking_band: row.ranking_band || RANKING_OVERRIDES[row.university] || "Indicative only",
      tuition_estimate: row.tuition_usd ? `$${row.tuition_usd}/year` : "Ask counselor",
      living_cost_estimate: livingEstimate(student.lifestyle_preference),
      estimated_total_yearly_cost: row.tuition_usd
        ? `Approx. $${row.tuition_usd} tuition + living depending on lifestyle`
        : "Depends on tuition + living style",
      scholarship_positioning: safeLower(row.scholarship_tag).includes("yes")
        ? "Scholarship/discount may be possible depending on profile and university policy."
        : "Scholarship may still be possible depending on profile and university policy.",
      why_this_fits: explainWhyThisFits(row, student),
      notes: [
        "Undergraduate Türkiye intake is Fall only.",
        row.tuition_basis === "Annualized Estimate"
          ? "Tuition shown here is an annualized estimate converted from a whole-program fee."
          : "Tuition shown here is treated as an annual fee.",
        "Seat reservation deposits are commonly around $1,000, depending on the university."
      ],
      diagnostics: row._diagnostics
    }))
  };
}

function resolveOrigin(requestOrigin) {
  if (ALLOWED_ORIGINS.includes(requestOrigin)) return requestOrigin;
  return ALLOWED_ORIGINS[0];
}

function parseJsonBody(req) {
  return new Promise((resolve, reject) => {
    let raw = "";
    req.on("data", (chunk) => {
      raw += chunk;
    });
    req.on("end", () => {
      try {
        resolve(raw ? JSON.parse(raw) : {});
      } catch (error) {
        reject(new Error("Invalid JSON request body"));
      }
    });
    req.on("error", reject);
  });
}

export default async function handler(req, res) {
  const origin = resolveOrigin(req.headers.origin);
  setCors(res, origin);

  if (req.method === "OPTIONS") {
    return res.status(204).end();
  }

  if (req.method === "GET") {
    return res.status(200).json({
      ok: true,
      country_scope: "Turkey only",
      loaded_program_rows: knowledge.programs.length,
      allowed_origins: ALLOWED_ORIGINS
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
      body = await parseJsonBody(req);
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

    const { rows: bestRows, diagnostics } = scorePrograms(merged);

    if (!bestRows.length) {
      return res.status(200).json({
        mode: "recommendations",
        country_scope: "Turkey only",
        docs_used: files.length > 0,
        diagnostics,
        summary: {
          budget_comment: "No strong match was found with the current inputs.",
          ranking_comment: "",
          scholarship_comment: "",
          lifestyle_comment: "",
          intake_comment: "For undergraduate Türkiye admissions, the intake is Fall only."
        },
        recommendations: [],
        message: "No confident program match was found. This response intentionally avoids returning unrelated programs."
      });
    }

    return res.status(200).json(buildRecommendationResponse(bestRows, merged, files.length > 0, diagnostics));
  } catch (error) {
    return res.status(500).json({
      error: error.message || "Server error"
    });
  }
}
