import fs from "fs";
import path from "path";
import Busboy from "busboy";
import pdfParse from "pdf-parse";
import mammoth from "mammoth";

export const config = {
  api: {
    bodyParser: false,
  },
};

const DEFAULT_ALLOWED_ORIGINS = [
  "https://tristar-education.myshopify.com",
  "https://tsapply.online",
  "https://www.tsapply.online",
];

const HUBSPOT_BASE = "https://api.hubapi.com";
const OPENAI_BASE = "https://api.openai.com/v1";
const OPENAI_MODEL = process.env.OPENAI_MODEL || "gpt-5.4";
const SOURCE_NEW_VALUE = "ai-recommender";
const NOTE_TO_CONTACT_ASSOCIATION_TYPE_ID = 202;

const HUBSPOT_PROP = {
  firstname: "firstname",
  lastname: "lastname",
  email: "email",
  phone: "phone",
  date_of_birth: "date_of_birth",
  gender: "gender",
  citizenship_country: "country",
  address_country: "country_2",
  education_system: "current_qualification",
  intended_program_single_field: "majors",
  semester_interested: "semester_interested_in",
  budget_year: "budget_year",
  source_new: "source_new",
  passport_first_page: "passport_url",
  student_picture: "student_picture_url",
  high_school_transcripts: "transcripts_url",
  entrance_exam: "entrance_exam_url",
  english_exam: "english_exam_url",
  personal_statement: "personal_statement_url",
  supporting_documents: "supporting_docs_urls",
  grades_profile_fallback: "comments",
};

const CATEGORY_BASE_SCORES = {
  "Top Tier": 40,
  "Strong Mid-Tier": 30,
  "Budget-Friendly": 20,
  General: 10,
  "Extreme Budget": 5,
  "Non-Preferred": -10,
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
  "atlas university": "Extreme Budget",
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
  "TED University": "Limited ranking presence",
};

const PROGRAM_FAMILIES = [
  {
    key: "business_administration",
    label: "Business Administration",
    aliases: ["business administration", "bba", "business", "management", "business management"],
    positivePatterns: [/business administration/, /\bbba\b/],
    negativePatterns: [
      /management engineering/,
      /aviation management/,
      /sports management/,
      /logistics management/,
      /management information systems/,
    ],
  },
  {
    key: "economics",
    label: "Economics",
    aliases: ["economics", "economics and finance", "economics & finance", "finance", "banking"],
    positivePatterns: [
      /economics/,
      /economics\s*(and|&)\s*finance/,
      /finance and banking/,
      /finance\s*&\s*banking/,
      /international trade and finance/,
      /international trade\s*&\s*finance/,
    ],
    negativePatterns: [/computer/, /software/, /data science/, /artificial intelligence/],
  },
  {
    key: "computer",
    label: "Computer Engineering",
    aliases: ["computer science", "cs", "computer engineering", "software engineering", "software development"],
    positivePatterns: [/computer science/, /computer engineering/, /software engineering/, /software development/],
    negativePatterns: [/economics/, /finance/, /psychology/, /history/, /comparative literature/],
  },
  {
    key: "electrical",
    label: "Electrical & Electronics Engineering",
    aliases: ["electrical", "electrical engineering", "electronics engineering", "electrical & electronics", "electrical-electronics"],
    positivePatterns: [/electrical/, /electronics engineering/, /electrical\s*&\s*electronics/, /electrical-?electronics/],
    negativePatterns: [],
  },
  {
    key: "psychology",
    label: "Psychology",
    aliases: ["psychology"],
    positivePatterns: [/psychology/, /psychological counselling/],
    negativePatterns: [],
  },
  {
    key: "medicine",
    label: "Medicine",
    aliases: ["medicine", "md", "doctor of medicine"],
    positivePatterns: [/medicine/, /doctor of medicine/, /\bmd\b/],
    negativePatterns: [/dentistry/, /nursing/, /pharmacy/],
  },
  {
    key: "dentistry",
    label: "Dentistry",
    aliases: ["dentistry", "dental", "bds"],
    positivePatterns: [/dentistry/, /dental/, /\bbds\b/],
    negativePatterns: [/medicine/, /nursing/],
  },
  {
    key: "nursing",
    label: "Nursing",
    aliases: ["nursing"],
    positivePatterns: [/nursing/],
    negativePatterns: [],
  },
  {
    key: "architecture",
    label: "Architecture",
    aliases: ["architecture", "interior architecture"],
    positivePatterns: [/architecture/],
    negativePatterns: [/archaeology/],
  },
  {
    key: "artificial_intelligence",
    label: "Artificial Intelligence",
    aliases: ["artificial intelligence", "ai", "machine learning"],
    positivePatterns: [/artificial intelligence/, /ai engineering/, /ai and data/],
    negativePatterns: [/economics/, /psychology/],
  },
  {
    key: "data_science",
    label: "Data Science",
    aliases: ["data science", "data analytics", "analytics", "data engineering"],
    positivePatterns: [/data science/, /analytics/, /data engineering/],
    negativePatterns: [/economics/, /finance/, /computer engineering/],
  },
];

function getAllowedOrigins() {
  const envOrigins = String(process.env.ALLOWED_ORIGINS || process.env.FRONTEND_ORIGIN || "")
    .split(",")
    .map((x) => x.trim())
    .filter(Boolean);

  return [...new Set([...DEFAULT_ALLOWED_ORIGINS, ...envOrigins])];
}

function setCors(res, origin) {
  if (origin) {
    res.setHeader("Access-Control-Allow-Origin", origin);
    res.setHeader("Vary", "Origin");
  }
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Max-Age", "86400");
}

function resolveOrigin(requestOrigin) {
  const allowedOrigins = getAllowedOrigins();
  if (!requestOrigin) return allowedOrigins[0] || null;
  return allowedOrigins.includes(requestOrigin) ? requestOrigin : null;
}

function assertServerConfig() {
  const missing = [];
  if (!process.env.OPENAI_API_KEY) missing.push("OPENAI_API_KEY");
  if (!process.env.HUBSPOT_PRIVATE_APP_TOKEN && !process.env.HUBSPOT_TOKEN) {
    missing.push("HUBSPOT_PRIVATE_APP_TOKEN");
  }
  if (!OPENAI_MODEL) missing.push("OPENAI_MODEL");

  if (missing.length) {
    throw new Error(`Missing required environment variables: ${missing.join(", ")}`);
  }
}

function normalizeText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function normalizeGender(value) {
  const v = normalizeText(value).toLowerCase();

  if (v === "male") return "Male";
  if (v === "female") return "Female";

  return "";
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

function uniqueNonEmpty(arr) {
  return [...new Set(arr.map((x) => normalizeText(x)).filter(Boolean))];
}

function parseMultipart(req) {
  return new Promise((resolve, reject) => {
    const bb = Busboy({ headers: req.headers });
    const fields = {};
    const files = [];

    bb.on("field", (name, val) => {
      if (fields[name] === undefined) fields[name] = val;
      else if (Array.isArray(fields[name])) fields[name].push(val);
      else fields[name] = [fields[name], val];
    });

    bb.on("file", (name, file, info) => {
      const chunks = [];

      file.on("data", (data) => chunks.push(data));

      file.on("end", () => {
        if (!info.filename) return;
        files.push({
          fieldname: name,
          filename: info.filename,
          mimeType: info.mimeType || "application/octet-stream",
          buffer: Buffer.concat(chunks),
        });
      });

      file.on("error", reject);
    });

    bb.on("finish", () => resolve({ fields, files }));
    bb.on("error", reject);

    req.pipe(bb);
  });
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
      } catch {
        reject(new Error("Invalid JSON request body"));
      }
    });

    req.on("error", reject);
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

function fileToDataUrl(file) {
  return `data:${file.mimeType};base64,${file.buffer.toString("base64")}`;
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

function detectRequestedDegree(input) {
  const q = normalizeForMatch(input);
  return {
    wants_ba: /\bba\b|\bbachelor of arts\b/.test(q),
    wants_bs: /\bbs\b|\bbsc\b|\bbachelor of science\b/.test(q),
    wants_bba: /\bbba\b|\bbachelor of business administration\b/.test(q),
  };
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
    familyLabel: family?.label || raw,
    ...degree,
  };
}

function getProgramSearchBlob(row) {
  return normalizeForMatch([
    row.program_title,
    row.major_tag,
    row.speciality,
    row.raw_title,
    row.raw_tags,
  ].join(" "));
}

function titleDegreeMatch(programTitle, intent) {
  const title = normalizeForMatch(programTitle);
  if (intent.wants_bba) return /business administration|\bbba\b/.test(title);
  if (intent.wants_bs) return /(\bbsc\b|bachelor of science|bachelors of science)/.test(title);
  if (intent.wants_ba) return /(\bba\b|bachelor of arts|bachelors of arts)/.test(title);
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
  const queryTokens = tokenize(intendedProgram).filter(
    (token) => token.length >= 3 && !["bachelor", "arts", "science"].includes(token)
  );

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

function categorizeUniversity(programRow) {
  if (programRow.category_tier) return programRow.category_tier;
  return UNIVERSITY_CATEGORY_MAP[safeLower(programRow.university)] || "General";
}

function categoryBaseScore(category) {
  return CATEGORY_BASE_SCORES[category] ?? 0;
}

function rankingScore(university, rankingImportance) {
  if (safeLower(rankingImportance) !== "yes") return 0;

  const category =
    typeof university === "string"
      ? UNIVERSITY_CATEGORY_MAP[safeLower(university)] || "General"
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
    if (["sabanci university", "ozyegin university", "koc university", "isik university"].includes(university)) {
      return 8;
    }
    return 0;
  }

  if (pref.includes("city")) {
    if (campusStyle.includes("city")) return 8;
    if (
      [
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
        "istinye university",
      ].includes(university)
    ) {
      return 8;
    }
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

function buildMajorsField(intendedProgram, selectedPrograms) {
  const selected = uniqueNonEmpty(selectedPrograms);
  let text = `Intended: ${normalizeText(intendedProgram) || ""}`.trim();

  if (selected.length) {
    text += "\n\nSelected: " + selected.join(" | ");
  }

  return text;
}

function parseMajorsField(value) {
  const raw = String(value || "").trim();
  if (!raw) return { intended: "", selected: [] };

  const intendedMatch = raw.match(/^Intended:\s*(.*?)(?:\n\nSelected:\s*(.*))?$/s);

  if (intendedMatch) {
    const intended = normalizeText(intendedMatch[1] || "");
    const selectedRaw = normalizeText(intendedMatch[2] || "");
    const selected = selectedRaw
      ? selectedRaw.split("|").map((item) => normalizeText(item)).filter(Boolean)
      : [];

    return { intended, selected: uniqueNonEmpty(selected) };
  }

  return { intended: normalizeText(raw), selected: [] };
}

function selectionKey(program, university) {
  return normalizeText([program, university].filter(Boolean).join(" - "));
}

function livingEstimate(lifestylePreference) {
  const pref = safeLower(lifestylePreference);
  if (pref.includes("campus")) return "$4,400 to $8,000";
  if (pref.includes("city")) return "$6,400 to $10,000";
  return "$4,400 to $10,000";
}

function explainWhyThisFits(row, student) {
  const reasons = [];
  const intent = parseProgramIntent(student.intended_program || "");

  if (intent.familyLabel) reasons.push(`Matched to intended program family: ${intent.familyLabel}.`);
  reasons.push(`Selected within the ${row._category} category based on TriStar counseling positioning.`);
  if (student.yearly_budget_total) {
    reasons.push(`Compared against the provided yearly budget of ${student.yearly_budget_total}.`);
  }
  if (student.lifestyle_preference) {
    reasons.push(`Considered lifestyle preference: ${student.lifestyle_preference}.`);
  }

  return reasons;
}

function buildRecommendationResponse(bestRows, student, selectedPrograms) {
  return {
    mode: "recommendations",
    country_scope: "Turkey only",
    selected_programs: uniqueNonEmpty(selectedPrograms),
    summary: {
      budget_comment: `Recommendations were matched against an approximate yearly tuition + living budget of ${
        student.yearly_budget_total || "not fully specified"
      }.`,
      ranking_comment:
        safeLower(student.ranking_importance) === "yes"
          ? "Ranking-sensitive options were prioritized."
          : "Ranking was not treated as the primary driver.",
      scholarship_comment: "Scholarships are shown as possible/estimated, not guaranteed.",
      lifestyle_comment: student.lifestyle_preference
        ? `Lifestyle preference considered: ${student.lifestyle_preference}.`
        : "Lifestyle preference was not fully specified.",
      intake_comment: "For undergraduate Türkiye admissions, the intake is Fall only.",
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
        row.tuition_basis === "Annualized Estimate" || row.tuition_basis === "Converted from Entire Program"
          ? "Tuition shown here is an annualized estimate converted from a whole-program fee."
          : "Tuition shown here is treated as an annual fee.",
        "Seat reservation deposits are commonly around $1,000, depending on the university.",
      ],
    })),
  };
}

function loadPrograms() {
  const datasetPath = path.join(process.cwd(), "knowledge", "Programs.json");
  if (!fs.existsSync(datasetPath)) {
    throw new Error("Programs dataset not found at knowledge/Programs.json");
  }

  const raw = JSON.parse(fs.readFileSync(datasetPath, "utf8"));

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
      raw_tags: normalizeText(row.Source_Raw_Tags || row.Raw_Tags),
    }))
    .filter((row) => row.university && row.program_title && safeLower(row.degree_level).includes("bachelor"));
}

const PROGRAMS = loadPrograms();

function scorePrograms(student) {
  const intent = parseProgramIntent(student.intended_program || "");

  const strictPool = PROGRAMS.filter((row) => strictProgramFilter(row, student.intended_program || ""));
  let pool = strictPool.length
    ? strictPool
    : PROGRAMS.filter((row) => rowMatchesFamily(row, intent.family) && titleDegreeMatch(row.program_title, intent));

  if (!pool.length) pool = PROGRAMS;

  const scored = pool
    .map((row) => {
      const category = categorizeUniversity(row);
      const total = [
        categoryBaseScore(category),
        programMatchScore(row, student.intended_program),
        rankingScore(row.university, student.ranking_importance),
        scholarshipScore(row, student.scholarship_preference),
        lifestyleScore(row, student.lifestyle_preference),
        budgetScore(row, student.yearly_budget_total, student.lifestyle_preference),
      ].reduce((sum, n) => sum + n, 0);

      return { ...row, _category: category, _score: total };
    })
    .filter((row) => row._score > -500)
    .sort(
      (a, b) =>
        b._score - a._score ||
        (a.tuition_usd ?? Infinity) - (b.tuition_usd ?? Infinity) ||
        a.university.localeCompare(b.university)
    );

  const finalRows = [];
  const seen = new Set();

  for (const row of scored) {
    if (finalRows.length >= 12) break;
    if (row._score < 20) continue;

    const key = `${safeLower(row.university)}::${safeLower(row.program_title)}`;
    if (seen.has(key)) continue;

    seen.add(key);
    finalRows.push(row);
  }

  return finalRows;
}

async function hubspotRequest(pathname, options = {}) {
  const token = process.env.HUBSPOT_PRIVATE_APP_TOKEN || process.env.HUBSPOT_TOKEN;
  if (!token) throw new Error("Missing HubSpot token");

  const response = await fetch(`${HUBSPOT_BASE}${pathname}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(options.headers || {}),
    },
  });

  const text = await response.text();
  
  let data = {};

  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    data = { raw: text };
  }

  if (!response.ok) {
    console.error("HubSpot request failed:", data?.message || "Unknown error");

    const detailedMessage =
      data?.message ||
      data?.error ||
      data?.errors?.map((e) => {
        const field = e?.context?.propertyName || e?.name || "unknown_field";
        const invalidValue =
          e?.context?.value ||
          e?.context?.invalidValue ||
          e?.in ||
          "unknown_value";
        return `${field}: ${invalidValue}`;
      }).join(" | ") ||
      `HubSpot request failed: ${response.status}`;

    throw new Error(detailedMessage);
  }

  return data;
}

async function findContactByEmail(email, extraProps = []) {
  const normalizedEmail = normalizeText(email).toLowerCase();
  if (!normalizedEmail) return null;

  const properties = uniqueNonEmpty([
    HUBSPOT_PROP.firstname,
    HUBSPOT_PROP.lastname,
    HUBSPOT_PROP.email,
    HUBSPOT_PROP.phone,
    HUBSPOT_PROP.intended_program_single_field,
    ...extraProps,
  ]);

  const result = await hubspotRequest("/crm/v3/objects/contacts/search", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      filterGroups: [
        {
          filters: [{ propertyName: "email", operator: "EQ", value: normalizedEmail }],
        },
      ],
      properties,
      limit: 1,
    }),
  });

  return result.results?.[0] || null;
}

async function createOrUpdateContactByEmail(properties) {
  const email = normalizeText(properties[HUBSPOT_PROP.email] || "").toLowerCase();
  if (!email) throw new Error("Email is required");

  const existing = await findContactByEmail(email, Object.values(HUBSPOT_PROP));

  if (existing?.id) {
    await hubspotRequest(`/crm/v3/objects/contacts/${existing.id}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ properties }),
    });

    return { id: existing.id, properties: existing.properties || {}, created: false };
  }

  const created = await hubspotRequest("/crm/v3/objects/contacts", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ properties }),
  });

  return { id: created.id, properties: created.properties || {}, created: true };
}

async function uploadFileToHubSpot(file, formFields) {
  const token = process.env.HUBSPOT_PRIVATE_APP_TOKEN || process.env.HUBSPOT_TOKEN;
  if (!token) throw new Error("Missing HubSpot token");

  const cleanName = buildCleanFileName(file.fieldname, file.filename, formFields);

  const uploadForm = new FormData();
  uploadForm.append("file", new Blob([file.buffer], { type: file.mimeType }), cleanName);
  uploadForm.append("fileName", cleanName);
  uploadForm.append("folderPath", "/ai-recommender");
  uploadForm.append("options", JSON.stringify({ access: "PUBLIC_NOT_INDEXABLE" }));

  const response = await fetch(`${HUBSPOT_BASE}/files/v3/files`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}` },
    body: uploadForm,
  });

  const text = await response.text();
  let data = {};

  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    data = { raw: text };
  }

  if (!response.ok) {
    throw new Error(data.message || `HubSpot file upload failed: ${response.status}`);
  }

  const fileId = data.id;
  if (!fileId) {
    throw new Error("HubSpot did not return a file ID: " + JSON.stringify(data));
  }

  return {
    id: String(fileId),
    url: data.url || data.defaultHostingUrl || "",
    name: cleanName,
    fieldname: file.fieldname,
  };
}

function setIfPresent(obj, key, value) {
  const normalized = normalizeText(value);
  if (normalized) obj[key] = normalized;
}

function buildContactPropertiesFromForm(form, existingMajorsValue = "") {
  const parsedMajors = parseMajorsField(existingMajorsValue);
  const intendedProgram = normalizeText(form.intended_program || parsedMajors.intended);
  const properties = {};

  setIfPresent(properties, HUBSPOT_PROP.firstname, form.first_name);
  setIfPresent(properties, HUBSPOT_PROP.lastname, form.last_name);
  setIfPresent(properties, HUBSPOT_PROP.email, normalizeText(form.email).toLowerCase());
  setIfPresent(properties, HUBSPOT_PROP.phone, form.phone);
  setIfPresent(properties, HUBSPOT_PROP.date_of_birth, form.date_of_birth);
  setIfPresent(properties, HUBSPOT_PROP.gender, normalizeGender(form.gender));
  setIfPresent(properties, HUBSPOT_PROP.citizenship_country, form.citizenship_country);
  setIfPresent(properties, HUBSPOT_PROP.address_country, form.address_country);
  setIfPresent(properties, HUBSPOT_PROP.education_system, form.education_system);
  setIfPresent(properties, HUBSPOT_PROP.semester_interested, form.semester_interested);
  setIfPresent(properties, HUBSPOT_PROP.budget_year, form.yearly_budget_total);
  setIfPresent(properties, HUBSPOT_PROP.source_new, SOURCE_NEW_VALUE);
  setIfPresent(properties, HUBSPOT_PROP.grades_profile_fallback, form.grades_profile);

  properties[HUBSPOT_PROP.intended_program_single_field] = buildMajorsField(
    intendedProgram,
    parsedMajors.selected
  );

  return properties;
}

async function createHubSpotNoteForContact(contactId, bodyText, attachmentIds = []) {
  const cleanAttachmentIds = attachmentIds
    .map((id) => normalizeText(id))
    .filter(Boolean);

  const properties = {
    hs_timestamp: new Date().toISOString(),
    hs_note_body: bodyText,
  };

  if (cleanAttachmentIds.length) {
    properties.hs_attachment_ids = cleanAttachmentIds.join(";");
  }

  await hubspotRequest("/crm/v3/objects/notes", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      properties,
      associations: [
        {
          to: { id: contactId },
          types: [
            {
              associationCategory: "HUBSPOT_DEFINED",
              associationTypeId: NOTE_TO_CONTACT_ASSOCIATION_TYPE_ID,
            },
          ],
        },
      ],
    }),
  });
}

async function openaiRequest(body) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) throw new Error("Missing OpenAI API key");

  const response = await fetch(`${OPENAI_BASE}/responses`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  const text = await response.text();
  let data = {};

  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    data = { raw: text };
  }

  if (!response.ok) {
    throw new Error(data.error?.message || `OpenAI request failed: ${response.status}`);
  }

  return data;
}

function extractResponseText(responseJson) {
  if (typeof responseJson.output_text === "string" && responseJson.output_text) {
    return responseJson.output_text;
  }

  const parts = [];

  for (const item of responseJson.output || []) {
    for (const content of item.content || []) {
      if (content.type === "output_text" && content.text) parts.push(content.text);
    }
  }

  return parts.join("\n").trim();
}

function parseDateToIso(value) {
  const raw = normalizeText(value);
  if (!raw) return "";

  const cleaned = raw.replace(/[,]/g, " ").replace(/\s+/g, " ").trim();

  const direct = new Date(cleaned);
  if (!Number.isNaN(direct.getTime())) {
    const y = direct.getFullYear();
    const m = String(direct.getMonth() + 1).padStart(2, "0");
    const d = String(direct.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  const parts = cleaned.match(/(\d{1,4})[\/.\-\s](\d{1,2})[\/.\-\s](\d{1,4})/);
  if (!parts) return "";

  let a = Number(parts[1]);
  let b = Number(parts[2]);
  let c = Number(parts[3]);

  let year;
  let month;
  let day;

  if (String(a).length === 4) {
    year = a;
    month = b;
    day = c;
  } else if (String(c).length === 4) {
    year = c;

    // Prefer day-month-year for passport-style dates, but tolerate month-day-year
    if (a > 12) {
      day = a;
      month = b;
    } else if (b > 12) {
      day = b;
      month = a;
    } else {
      day = a;
      month = b;
    }
  } else {
    return "";
  }

  if (!year || !month || !day) return "";
  if (month < 1 || month > 12 || day < 1 || day > 31) return "";

  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function normalizePersonName(value) {
  return normalizeText(value)
    .toLowerCase()
    .replace(/[^a-z\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function tokenizePersonName(value) {
  return normalizePersonName(value).split(" ").filter(Boolean);
}

function namesMatchLoosely(formFullName, passportFullName) {
  const formTokens = tokenizePersonName(formFullName);
  const passportTokens = tokenizePersonName(passportFullName);

  if (!formTokens.length || !passportTokens.length) return false;

  const formSet = new Set(formTokens);
  const passportSet = new Set(passportTokens);

  const overlap = formTokens.filter((token) => passportSet.has(token)).length;
  const surnameMatch =
    formTokens.length && passportTokens.length
      ? formTokens[formTokens.length - 1] === passportTokens[passportTokens.length - 1]
      : false;

  const overlapRatio = overlap / formTokens.length;

  if (surnameMatch && overlapRatio >= 0.5) return true;
  if (overlapRatio >= 0.75) return true;

  const formJoined = formTokens.join(" ");
  const passportJoined = passportTokens.join(" ");

  if (passportJoined.includes(formJoined) || formJoined.includes(passportJoined)) return true;

  return false;
}

async function runPassportIdentityCheck(form, passportFile) {
  if (!passportFile) {
    return {
      passed: false,
      reason: "Passport file is missing.",
      passport_name: "",
      passport_date_of_birth: "",
    };
  }

  const input = [
    {
      role: "system",
      content: [
        {
          type: "input_text",
          text: [
            "You extract identity details from a passport image or passport PDF text.",
            "Return strict JSON only.",
            JSON.stringify(
              {
                passport_full_name: "",
                date_of_birth: "",
                confidence: "HIGH | MEDIUM | LOW",
                notes: "",
              },
              null,
              2
            ),
          ].join("\n"),
        },
      ],
    },
  ];

  const extractedText = await extractTextFromFile(passportFile);

  if (extractedText) {
    input.push({
      role: "user",
      content: [
        {
          type: "input_text",
          text: `Passport text:\n${extractedText.slice(0, 12000)}`,
        },
      ],
    });
  } else if (passportFile.mimeType.startsWith("image/")) {
    input.push({
      role: "user",
      content: [
        { type: "input_text", text: "Extract the passport holder full name and date of birth from this passport image." },
        { type: "input_image", image_url: fileToDataUrl(passportFile), detail: "high" },
      ],
    });
  } else {
    return {
      passed: false,
      reason: "Passport could not be read.",
      passport_name: "",
      passport_date_of_birth: "",
    };
  }

  const response = await openaiRequest({
    model: OPENAI_MODEL,
    input,
    text: { format: { type: "json_object" } },
  });

  const parsed = JSON.parse(extractResponseText(response) || "{}");

  const passportName = normalizeText(parsed.passport_full_name);
  const passportDob = parseDateToIso(parsed.date_of_birth);
  const formFullName = normalizeText([form.first_name, form.last_name].filter(Boolean).join(" "));
  const formDob = parseDateToIso(form.date_of_birth);

  const nameOk = namesMatchLoosely(formFullName, passportName);
  const dobOk = !!formDob && !!passportDob && formDob === passportDob;

  return {
    passed: nameOk && dobOk,
    name_ok: nameOk,
    dob_ok: dobOk,
    passport_name: passportName,
    passport_date_of_birth: passportDob,
    form_name: formFullName,
    form_date_of_birth: formDob,
    confidence: normalizeText(parsed.confidence) || "LOW",
    notes: normalizeText(parsed.notes),
    reason:
      nameOk && dobOk
        ? ""
        : !nameOk && !dobOk
          ? "Passport name and date of birth do not match the submitted form."
          : !nameOk
            ? "Passport name does not match the submitted form."
            : "Passport date of birth does not match the submitted form.",
  };
}

async function runDocumentVerification(form, files) {
  if (!files.length) {
    return {
      status: "NO_DOCUMENTS",
      confidence: "LOW",
      matched_fields: [],
      mismatched_fields: [],
      missing_or_unreadable: ["No files uploaded for AI review."],
      advisor_summary: "No documents were available for AI cross-checking.",
    };
  }

  const input = [
    {
      role: "system",
      content: [
        {
          type: "input_text",
          text: [
            "You are an admissions document cross-checking assistant.",
            "Compare the student's submitted form details against uploaded documents and compare the documents against each other.",
            "Do not claim authenticity.",
            "Do not block recommendations.",
            "Return strict JSON with this schema:",
            JSON.stringify(
              {
                status: "PASSED | FLAGGED | PARTIAL",
                confidence: "HIGH | MEDIUM | LOW",
                matched_fields: ["..."],
                mismatched_fields: ["..."],
                missing_or_unreadable: ["..."],
                advisor_summary: "...",
              },
              null,
              2
            ),
          ].join("\n"),
        },
      ],
    },
    {
      role: "user",
      content: [
        {
          type: "input_text",
          text: `Student form data:\n${JSON.stringify(form, null, 2)}`,
        },
      ],
    },
  ];

  for (const file of files.slice(0, 10)) {
    const localText = await extractTextFromFile(file);

    if (localText) {
      input.push({
        role: "user",
        content: [
          {
            type: "input_text",
            text: `Extracted text from ${file.filename}:\n${localText.slice(0, 12000)}`,
          },
        ],
      });
    } else if (file.mimeType.startsWith("image/")) {
      input.push({
        role: "user",
        content: [
          { type: "input_text", text: `Review this uploaded image file: ${file.filename}` },
          { type: "input_image", image_url: fileToDataUrl(file), detail: "high" },
        ],
      });
    }
  }

  const response = await openaiRequest({
    model: OPENAI_MODEL,
    input,
    text: { format: { type: "json_object" } },
  });

  const parsed = JSON.parse(extractResponseText(response) || "{}");

  return {
    status: normalizeText(parsed.status) || "PARTIAL",
    confidence: normalizeText(parsed.confidence) || "LOW",
    matched_fields: Array.isArray(parsed.matched_fields) ? parsed.matched_fields : [],
    mismatched_fields: Array.isArray(parsed.mismatched_fields) ? parsed.mismatched_fields : [],
    missing_or_unreadable: Array.isArray(parsed.missing_or_unreadable) ? parsed.missing_or_unreadable : [],
    advisor_summary: normalizeText(parsed.advisor_summary),
  };
}

function buildAdvisorNote(form, uploadedFiles, verification) {
  const lines = [
    "AI Document Cross-Check Summary",
    "",
    `Status: ${verification.status}`,
    `Confidence: ${verification.confidence}`,
    `Student: ${normalizeText([form.first_name, form.last_name].filter(Boolean).join(" ")) || "Unknown"}`,
    `Email: ${normalizeText(form.email)}`,
    `Intended Program: ${normalizeText(form.intended_program)}`,
    "",
    "Files received:",
    ...uploadedFiles.map((f) => `- ${f.fieldname}: ${f.name}${f.url ? ` (${f.url})` : ""}`),
    "",
    "Matched fields:",
    ...(verification.matched_fields.length ? verification.matched_fields.map((x) => `- ${x}`) : ["- None recorded"]),
    "",
    "Mismatched fields:",
    ...(verification.mismatched_fields.length
      ? verification.mismatched_fields.map((x) => `- ${x}`)
      : ["- None recorded"]),
    "",
    "Missing / unreadable:",
    ...(verification.missing_or_unreadable.length
      ? verification.missing_or_unreadable.map((x) => `- ${x}`)
      : ["- None recorded"]),
    "",
    "Advisor summary:",
    verification.advisor_summary || "No summary returned.",
  ];

  return lines.join("\n");
}

function groupFilesByField(files) {
  const map = new Map();

  for (const file of files) {
    if (!map.has(file.fieldname)) map.set(file.fieldname, []);
    map.get(file.fieldname).push(file);
  }

  return map;
}

function validateRequiredFinalFields(form, groupedFiles) {
  const requiredFields = [
    "first_name",
    "last_name",
    "date_of_birth",
    "gender",
    "citizenship_country",
    "email",
    "phone",
    "address_country",
    "education_system",
    "intended_program",
    "semester_interested",
    "yearly_budget_total",
    "ranking_importance",
    "scholarship_preference",
    "lifestyle_preference",
    "grades_profile",
    "candidate_type",
  ];

  for (const key of requiredFields) {
    if (!normalizeText(form[key])) throw new Error(`Missing required field: ${key}`);
  }

  const requiredFileFields = ["passport_first_page", "student_picture", "high_school_transcripts"];

  for (const key of requiredFileFields) {
    if (!(groupedFiles.get(key) || []).length) {
      throw new Error(`Missing required file: ${key}`);
    }
  }
}

async function handleCaptureStepOne(body) {
  const properties = {};

  setIfPresent(properties, HUBSPOT_PROP.firstname, body.first_name);
  setIfPresent(properties, HUBSPOT_PROP.lastname, body.last_name);
  setIfPresent(properties, HUBSPOT_PROP.email, normalizeText(body.email).toLowerCase());
  setIfPresent(properties, HUBSPOT_PROP.phone, body.phone);
  setIfPresent(properties, HUBSPOT_PROP.date_of_birth, body.date_of_birth);
  setIfPresent(properties, HUBSPOT_PROP.gender, normalizeGender(body.gender));
  setIfPresent(properties, HUBSPOT_PROP.citizenship_country, body.citizenship_country);
  setIfPresent(properties, HUBSPOT_PROP.address_country, body.address_country);
  setIfPresent(properties, HUBSPOT_PROP.source_new, SOURCE_NEW_VALUE);

  const existing = await findContactByEmail(body.email, [HUBSPOT_PROP.intended_program_single_field]);
  const existingMajors = existing?.properties?.[HUBSPOT_PROP.intended_program_single_field] || "";
  const parsedExisting = parseMajorsField(existingMajors);

  properties[HUBSPOT_PROP.intended_program_single_field] = buildMajorsField(
    parsedExisting.intended || normalizeText(body.intended_program),
    parsedExisting.selected
  );

  const result = await createOrUpdateContactByEmail(properties);
  return { ok: true, mode: "capture_step_1", contact_id: result.id };
}

async function handleTogglePreferredProgram(body) {
  const email = normalizeText(body.email).toLowerCase();
  if (!email) throw new Error("Email is required");

  const contact = await findContactByEmail(email, [HUBSPOT_PROP.intended_program_single_field]);
  if (!contact?.id) throw new Error("No HubSpot contact found for this email");

  const current = parseMajorsField(contact.properties?.[HUBSPOT_PROP.intended_program_single_field] || "");
  const intended = current.intended || normalizeText(body.intended_program);
  const selectedKey = selectionKey(body.selected_program, body.selected_university);

  let selected = current.selected.slice();
  if (selected.includes(selectedKey)) selected = selected.filter((x) => x !== selectedKey);
  else selected = uniqueNonEmpty([...selected, selectedKey]);

  const nextValue = buildMajorsField(intended, selected);

  await hubspotRequest(`/crm/v3/objects/contacts/${contact.id}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      properties: {
        [HUBSPOT_PROP.intended_program_single_field]: nextValue,
        [HUBSPOT_PROP.source_new]: SOURCE_NEW_VALUE,
      },
    }),
  });

  return { ok: true, mode: "toggle_preferred_program", selected_programs: selected };
}

async function handleGetMoreRecommendations(body) {
  const email = normalizeText(body.email).toLowerCase();
  const intendedProgram = normalizeText(body.intended_program);

  if (!email) throw new Error("Email is required");
  if (!intendedProgram) throw new Error("Intended program is required");

  const contact = await findContactByEmail(email, [HUBSPOT_PROP.intended_program_single_field]);
  if (!contact?.id) throw new Error("No HubSpot contact found for this email");

  const current = parseMajorsField(contact.properties?.[HUBSPOT_PROP.intended_program_single_field] || "");
  const selected = current.selected || [];

  const nextMajorsValue = buildMajorsField(intendedProgram, selected);

  await hubspotRequest(`/crm/v3/objects/contacts/${contact.id}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      properties: {
        [HUBSPOT_PROP.intended_program_single_field]: nextMajorsValue,
        [HUBSPOT_PROP.source_new]: SOURCE_NEW_VALUE,
      },
    }),
  });

  const student = {
    intended_program: intendedProgram,
    ranking_importance: body.ranking_importance,
    scholarship_preference: body.scholarship_preference,
    lifestyle_preference: body.lifestyle_preference,
    yearly_budget_total: body.yearly_budget_total,
  };

  const recommendations = scorePrograms(student);
  return buildRecommendationResponse(recommendations, student, selected);
}

async function handleFinalSubmit(fields, files) {
  const groupedFiles = groupFilesByField(files);
  validateRequiredFinalFields(fields, groupedFiles);

  const passportFile = (groupedFiles.get("passport_first_page") || [])[0];
  const passportIdentityCheck = await runPassportIdentityCheck(fields, passportFile);

if (!passportIdentityCheck.passed) {
  throw new Error(
      "Your passport details do not match the name or date of birth entered in the form. Please review Step 1 and upload the correct passport."
  );
}

  const existing = await findContactByEmail(fields.email, [HUBSPOT_PROP.intended_program_single_field]);
  const existingMajors = existing?.properties?.[HUBSPOT_PROP.intended_program_single_field] || "";
  const contactProps = buildContactPropertiesFromForm(fields, existingMajors);
  const contact = await createOrUpdateContactByEmail(contactProps);
  const contactId = contact.id;

  const uploaded = [];
  for (const file of files) {
    const uploadedFile = await uploadFileToHubSpot(file, fields);
    uploaded.push(uploadedFile);
  }

  const patchProps = {
    [HUBSPOT_PROP.source_new]: SOURCE_NEW_VALUE,
  };

  const firstFileUrl = (field) => uploaded.find((x) => x.fieldname === field)?.url || "";

  patchProps[HUBSPOT_PROP.passport_first_page] = firstFileUrl("passport_first_page") || undefined;
  patchProps[HUBSPOT_PROP.student_picture] = firstFileUrl("student_picture") || undefined;
  patchProps[HUBSPOT_PROP.high_school_transcripts] =
    uploaded.find((x) => x.fieldname === "high_school_transcripts")?.url || undefined;

  patchProps[HUBSPOT_PROP.entrance_exam] = firstFileUrl("entrance_exam") || undefined;
  patchProps[HUBSPOT_PROP.english_exam] = firstFileUrl("english_exam") || undefined;
  patchProps[HUBSPOT_PROP.personal_statement] = firstFileUrl("personal_statement") || undefined;

  const supportFileUrls = uploaded
    .filter((x) => x.fieldname === "supporting_documents")
    .map((x) => x.url)
    .filter(Boolean);

  patchProps[HUBSPOT_PROP.supporting_documents] = supportFileUrls.join("\n") || undefined;

  await hubspotRequest(`/crm/v3/objects/contacts/${contactId}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      properties: Object.fromEntries(
        Object.entries(patchProps).filter(([, v]) => v !== undefined)
      ),
    }),
  });

  const verification = await runDocumentVerification(fields, files);
  const advisorNote = buildAdvisorNote(fields, uploaded, verification);

  await createHubSpotNoteForContact(
    contactId,
    advisorNote,
    uploaded.map((f) => f.id)
  );

  const latestContact = await findContactByEmail(fields.email, [HUBSPOT_PROP.intended_program_single_field]);
  const majorsParsed = parseMajorsField(
    latestContact?.properties?.[HUBSPOT_PROP.intended_program_single_field] ||
      contactProps[HUBSPOT_PROP.intended_program_single_field]
  );

  const student = {
    intended_program: majorsParsed.intended || fields.intended_program,
    ranking_importance: fields.ranking_importance,
    scholarship_preference: fields.scholarship_preference,
    lifestyle_preference: fields.lifestyle_preference,
    yearly_budget_total: fields.yearly_budget_total,
  };

  const recommendations = scorePrograms(student);
  return buildRecommendationResponse(recommendations, student, majorsParsed.selected);
}

function buildCleanFileName(fieldname, originalName, formFields) {
  const ext = originalName.includes(".")
    ? "." + originalName.split(".").pop().toLowerCase()
    : "";

  const firstName = normalizeText(formFields.first_name);
  const lastName = normalizeText(formFields.last_name);
  const fullName = [firstName, lastName].filter(Boolean).join(" ") || "Student";

  const FIELD_LABELS = {
    passport_first_page: "Passport",
    student_picture: "Student Photo",
    high_school_transcripts: "Transcripts",
    entrance_exam: "Entrance Exam",
    english_exam: "English Exam",
    personal_statement: "Personal Statement",
    supporting_documents: "Supporting Document",
  };

  const label = FIELD_LABELS[fieldname] || "Document";
  return `${fullName} - ${label}${ext}`;
}

function publicErrorMessage(error) {
  return String(error?.message || "Something went wrong. Please try again.");
}

export default async function handler(req, res) {
  const origin = resolveOrigin(req.headers.origin);
  setCors(res, origin);

  if (req.method === "OPTIONS") {
    if (!origin && req.headers.origin) {
      return res.status(403).json({ error: "Origin not allowed" });
    }
    return res.status(204).end();
  }

  if (req.headers.origin && !origin) {
    return res.status(403).json({ error: "Origin not allowed" });
  }

  if (req.method === "GET") {
    try {
      assertServerConfig();
      return res.status(200).json({
        ok: true,
        message: "AI recommender backend is running.",
        loaded_program_rows: PROGRAMS.length,
        allowed_origins: getAllowedOrigins(),
      });
    } catch (error) {
      return res.status(500).json({
        ok: false,
        error: publicErrorMessage(error),
      });
    }
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    assertServerConfig();

    const contentType = req.headers["content-type"] || "";

    if (contentType.includes("multipart/form-data")) {
      const { fields, files } = await parseMultipart(req);
      const action = normalizeText(fields.action || "final_submit");

      if (action === "final_submit") {
        const result = await handleFinalSubmit(fields, files);
        return res.status(200).json(result);
      }

      return res.status(400).json({ error: "Unsupported multipart action" });
    }

    const body = await parseJsonBody(req);
    const action = normalizeText(body.action);

    if (action === "capture_step_1") {
  const result = await handleCaptureStepOne(body);
  return res.status(200).json(result);
}

if (action === "toggle_preferred_program") {
  const result = await handleTogglePreferredProgram(body);
  return res.status(200).json(result);
}

if (action === "get_more_recommendations") {
  const result = await handleGetMoreRecommendations(body);
  return res.status(200).json(result);
}

return res.status(400).json({ error: "Unsupported JSON action" });
  } catch (error) {
    console.error("Backend error:", error?.message || error);

    const message = publicErrorMessage(error);
    const status =
      message.includes("Please complete") ||
      message.includes("Please upload") ||
      message.includes("invalid") ||
      message.includes("could not find your profile") ||
      message.includes("Origin not allowed")
        ? 400
        : 500;

    return res.status(status).json({
  error: status === 500
    ? "Something went wrong. Please try again."
    : message
});
  }
}