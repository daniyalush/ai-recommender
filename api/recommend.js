import fs from "fs";
import path from "path";
import xlsx from "xlsx";

function setCors(res, origin = "*") {
  res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Access-Control-Max-Age", "86400");
}

function extractJSON(str) {
  if (!str) return null;
  const match = String(str).match(/\{[\s\S]*\}/);
  return match ? match[0] : null;
}

function normalizeText(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .replace(/[^\w\s%+&().,-]/g, "")
    .trim();
}

function safeLower(value) {
  return normalizeText(value).toLowerCase();
}

function toNumber(value) {
  if (value === null || value === undefined) return null;
  const cleaned = String(value).replace(/[^0-9.]/g, "");
  const num = parseFloat(cleaned);
  return Number.isFinite(num) ? num : null;
}

function sheetToRows(filePath) {
  if (!fs.existsSync(filePath)) return [];
  const workbook = xlsx.readFile(filePath);
  const rows = [];

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const json = xlsx.utils.sheet_to_json(sheet, { defval: "" });
    json.forEach((row) => {
      rows.push({ ...row, __sheet: sheetName, __file: path.basename(filePath) });
    });
  });

  return rows;
}

function getFirst(row, keys) {
  for (const key of keys) {
    if (row[key] !== undefined && row[key] !== null && String(row[key]).trim() !== "") {
      return row[key];
    }
  }
  return "";
}

function mapProgramRows(rawRows) {
  return rawRows
    .map((row) => {
      const university = getFirst(row, [
        "University",
        "university",
        "UNIVERSITY",
        "Institution",
        "institution",
        "School",
        "school"
      ]);

      const program = getFirst(row, [
        "Program",
        "PROGRAM",
        "program",
        "Programme",
        "PROGRAMME",
        "Major",
        "major",
        "Department",
        "department"
      ]);

      const language = getFirst(row, [
        "Language",
        "language",
        "Medium",
        "medium",
        "Instruction Language"
      ]);

      const tuition = getFirst(row, [
        "Tuition",
        "tuition",
        "Tuition Fee",
        "TUITION FEE / PER YEAR",
        "Fee",
        "fee",
        "Annual Fee",
        "annual fee",
        "Price",
        "price"
      ]);

      const scholarship = getFirst(row, [
        "Scholarship",
        "scholarship",
        "Scholarship %",
        "Scholarship Percentage",
        "Discount",
        "discount"
      ]);

      const rankingBand = getFirst(row, [
        "Ranking",
        "ranking",
        "THE Ranking",
        "Ranking Band"
      ]);

      const campusStyle = getFirst(row, [
        "Campus Style",
        "campus style",
        "Lifestyle",
        "lifestyle",
        "Type"
      ]);

      const category = getFirst(row, [
        "Category",
        "category",
        "Tier",
        "tier"
      ]);

      if (!university && !program) return null;

      return {
        university: normalizeText(university),
        program: normalizeText(program),
        language: normalizeText(language),
        tuition: normalizeText(tuition),
        scholarship: normalizeText(scholarship),
        ranking_band: normalizeText(rankingBand),
        campus_style: normalizeText(campusStyle),
        category: normalizeText(category),
        source_file: row.__file || "",
        source_sheet: row.__sheet || ""
      };
    })
    .filter(Boolean);
}

function loadKnowledge() {
  const knowledgeDir = path.join(process.cwd(), "knowledge");

  const files = [
    "Aydin 2026-2027 Tuition Fees.xlsx",
    "Bilgi Fall 2026 Tuition Prices with Scholarship.xlsx",
    "Istanbul Okan University International Student 2026-2027 Price List (1).xlsx",
    "Istinye 2026-2027 Fall Undergraduate Programs.xlsx",
    "Programs_structured.xlsx"
  ];

  const raw = files.flatMap((file) => sheetToRows(path.join(knowledgeDir, file)));
  const programs = mapProgramRows(raw);

  return {
    programs,
    rules: {
      country_scope: "Turkey only",
      intake: "Fall only for undergraduate admissions",
      sat_policy: "SAT is not mandatory for most Turkish universities but may help scholarship chances",
      deposit: "Seat reservation deposit is usually around $1,000",
      living_costs: {
        dorm_total_year: "$4,400 to $8,000",
        shared_apartment_total_year: "$6,400 to $10,000",
        food_transport_month: "$300 to $500"
      },
      mandatory_questions: [
        "Do you have any specific university preference?",
        "Is university ranking an important factor for you?",
        "What is your approximate yearly budget for tuition and living?",
        "Are you specifically looking for high scholarships?",
        "Do you prefer a city-center university experience or a campus-based environment?",
        "What education system are you applying with?",
        "Do you have an SAT score? If yes, what score?",
        "What is your intended program/major?",
        "Are you in final year or already completed? Share grades/predicted.",
        "Are you a private candidate or applying through a school?"
      ]
    },
    categories: {
      top_tier: ["Sabanci University", "Koc University"],
      strong_mid_tier: [
        "Ozyegin University",
        "Kadir Has University",
        "Istanbul Bilgi University",
        "Bahcesehir University",
        "Istinye University"
      ],
      budget_friendly: [
        "Istanbul Medipol University",
        "Isik University",
        "TED University",
        "Yasar University",
        "Uskudar University",
        "Istanbul Aydin University",
        "Istanbul Okan University"
      ],
      extreme_budget: ["Beykoz University", "Kent University"]
    },
    ranking_overrides: {
      "Sabanci University": "THE 351–400 band",
      "Koc University": "THE 301–350 band",
      "Kadir Has University": "THE 601–800 band",
      "Ozyegin University": "THE 801–1000 band",
      "Bahcesehir University": "THE 801–1000 band",
      "Istanbul Medipol University": "THE 801–1000 band",
      "Istanbul Bilgi University": "THE 1200+ band",
      "TED University": "Limited ranking presence"
    }
  };
}

function categorizeUniversity(name, knowledge) {
  const n = safeLower(name);

  for (const item of knowledge.categories.top_tier) {
    if (safeLower(item) === n) return "Top Tier";
  }
  for (const item of knowledge.categories.strong_mid_tier) {
    if (safeLower(item) === n) return "Strong Mid-Tier";
  }
  for (const item of knowledge.categories.budget_friendly) {
    if (safeLower(item) === n) return "Budget-Friendly";
  }
  for (const item of knowledge.categories.extreme_budget) {
    if (safeLower(item) === n) return "Extreme Budget";
  }

  return "General";
}

function getCategoryPriority(category) {
  const map = {
    "Top Tier": 4,
    "Strong Mid-Tier": 3,
    "Budget-Friendly": 2,
    "Extreme Budget": 1,
    "General": 0
  };
  return map[category] ?? 0;
}

function estimateLiving(stylePref, knowledge) {
  const pref = safeLower(stylePref);
  if (pref.includes("campus")) {
    return knowledge.rules.living_costs.dorm_total_year;
  }
  if (pref.includes("city")) {
    return knowledge.rules.living_costs.shared_apartment_total_year;
  }
  return "$4,400 to $10,000";
}

function budgetFitScore(tuitionValue, budgetValue) {
  if (!budgetValue || !tuitionValue) return 0;
  if (tuitionValue <= budgetValue * 0.65) return 3;
  if (tuitionValue <= budgetValue * 0.85) return 2;
  if (tuitionValue <= budgetValue) return 1;
  return -2;
}

function rankingScore(university, rankingImportance, knowledge) {
  if (safeLower(rankingImportance) !== "yes") return 0;
  const category = categorizeUniversity(university, knowledge);
  return getCategoryPriority(category);
}

function lifestyleScore(university, stylePref) {
  const pref = safeLower(stylePref);
  const uni = safeLower(university);

  if (!pref) return 0;

  const campusBased = ["sabanci", "ozyegin", "isik"];
  const cityBased = ["bilgi", "kadir has", "bahcesehir", "medipol", "aydin", "okan", "uskudar", "istinye"];

  if (pref.includes("campus")) {
    return campusBased.some((x) => uni.includes(x)) ? 2 : 0;
  }

  if (pref.includes("city")) {
    return cityBased.some((x) => uni.includes(x)) ? 2 : 0;
  }

  return 0;
}

function scholarshipScore(row, scholarshipPreference) {
  const wants = safeLower(scholarshipPreference);
  if (!wants) return 0;

  const combined = safeLower(`${row.scholarship} ${row.tuition} ${row.program} ${row.university}`);
  if (wants.includes("yes") || wants.includes("high")) {
    if (
      combined.includes("75") ||
      combined.includes("60") ||
      combined.includes("50") ||
      combined.includes("scholar")
    ) {
      return 2;
    }
    return 0;
  }

  return 0;
}

function preferenceScore(row, universityPreference, intendedProgram) {
  let score = 0;

  const uniPref = safeLower(universityPreference);
  const intended = safeLower(intendedProgram);
  const uni = safeLower(row.university);
  const program = safeLower(row.program);

  if (uniPref && uni.includes(uniPref)) score += 4;
  if (intended && program.includes(intended)) score += 5;

  const aliases = {
    "computer science": ["computer engineering", "software engineering", "artificial intelligence", "management information systems", "computer programming"],
    "software engineering": ["software engineering", "computer engineering", "back-end software development"],
    "artificial intelligence": ["artificial intelligence", "robotics", "computer engineering"],
    "business": ["business administration", "management", "economics", "finance", "international trade"],
    "psychology": ["psychology", "applied psychology", "clinical psychology"],
    "medicine": ["medicine"],
    "dentistry": ["dentistry"],
    "nursing": ["nursing"],
    "architecture": ["architecture", "interior architecture", "urban design"],
    "media": ["media", "communication", "journalism", "radio", "television", "visual communication"]
  };

  Object.entries(aliases).forEach(([key, vals]) => {
    if (intended.includes(key) && vals.some((v) => program.includes(v))) {
      score += 3;
    }
  });

  return score;
}

function buildReasons(row, student, knowledge) {
  const reasons = [];
  const notes = [];

  const category = categorizeUniversity(row.university, knowledge);
  reasons.push(`Positioned in the ${category} category for Turkey counseling.`);

  if (student.intended_program) {
    reasons.push(`Matches the student's intended area: ${student.intended_program}.`);
  }

  if (student.lifestyle_preference) {
    reasons.push(`Aligned with the preferred lifestyle: ${student.lifestyle_preference}.`);
  }

  if (row.tuition) {
    reasons.push(`Tuition information is available from the uploaded fee sheets.`);
  }

  const ranking = knowledge.ranking_overrides[row.university] || row.ranking_band;
  if (ranking) {
    notes.push(`Ranking positioning: ${ranking}.`);
  }

  notes.push(`Undergraduate Turkey intake is Fall only.`);
  notes.push(`Scholarships should be positioned as estimated/possible ranges, not guaranteed.`);

  return { reasons: reasons.slice(0, 3), notes: notes.slice(0, 2) };
}

function scoreRows(programRows, student, knowledge) {
  const budget = toNumber(student.yearly_budget_total);

  const scored = programRows.map((row) => {
    const tuitionVal = toNumber(row.tuition);
    const category = categorizeUniversity(row.university, knowledge);

    let score = 0;
    score += preferenceScore(row, student.university_preference, student.intended_program);
    score += rankingScore(row.university, student.ranking_importance, knowledge);
    score += scholarshipScore(row, student.scholarship_preference);
    score += lifestyleScore(row.university, student.lifestyle_preference);
    score += budgetFitScore(tuitionVal, budget);
    score += getCategoryPriority(category);

    if (student.sat_score && safeLower(student.sat_score) !== "no") {
      score += 0.5;
    }

    return {
      ...row,
      _score: score,
      _category: category
    };
  });

  scored.sort((a, b) => b._score - a._score);

  const deduped = [];
  const seen = new Set();

  for (const row of scored) {
    const key = `${safeLower(row.university)}|${safeLower(row.program)}`;
    if (!seen.has(key)) {
      seen.add(key);
      deduped.push(row);
    }
  }

  return deduped.slice(0, 3);
}

function buildQuestionResponse(body, missing) {
  return {
    mode: "questions",
    country_scope: "Turkey only",
    message: "Please answer these before I recommend universities.",
    questions: missing.map((f) => ({
      key: f.key,
      question: f.label
    })),
    current_answers: body
  };
}

function buildRecommendationResponse(bestRows, student, knowledge) {
  const living = estimateLiving(student.lifestyle_preference, knowledge);

  return {
    mode: "recommendations",
    country_scope: "Turkey only",
    summary: {
      budget_comment: student.yearly_budget_total
        ? `Recommendations were filtered around an approximate yearly tuition + living budget of ${student.yearly_budget_total}.`
        : "Budget was not fully specified.",
      ranking_comment: safeLower(student.ranking_importance) === "yes"
        ? "Ranking-sensitive options were prioritized first."
        : "Ranking was not treated as the primary driver.",
      scholarship_comment: student.scholarship_preference
        ? "Scholarships are positioned as estimated possibilities, not guarantees."
        : "Scholarship preference was not treated as the main driver.",
      lifestyle_comment: student.lifestyle_preference
        ? `Lifestyle preference considered: ${student.lifestyle_preference}.`
        : "Lifestyle preference was not clearly specified.",
      intake_comment: "For undergraduate Turkey admissions, the relevant intake is Fall."
    },
    recommendations: bestRows.map((row) => {
      const rankingBand = knowledge.ranking_overrides[row.university] || row.ranking_band || "Indicative only";
      const { reasons, notes } = buildReasons(row, student, knowledge);

      return {
        university: row.university || "",
        program: row.program || "",
        category: row._category || "General",
        ranking_band: rankingBand,
        tuition_estimate: row.tuition || "Ask counselor",
        living_cost_estimate: living,
        estimated_total_yearly_cost: "Estimate depends on tuition + living style",
        scholarship_positioning: row.scholarship
          ? `Possible scholarship positioning based on source sheet: ${row.scholarship}`
          : "Scholarship may be possible depending on profile and university policy.",
        why_this_fits: reasons,
        notes
      };
    })
  };
}

const knowledgeCache = loadKnowledge();

export default async function handler(req, res) {
  const allowedOrigins = [
    "https://tristar-education.myshopify.com"
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
      loaded_program_rows: knowledgeCache.programs.length
    });
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const body = req.body || {};

    const profilingFields = [
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
    ];

    const missing = profilingFields.filter(
      (f) => !String(body[f.key] || "").trim()
    );

    if (missing.length > 0) {
      return res.status(200).json(buildQuestionResponse(body, missing));
    }

    const bestRows = scoreRows(knowledgeCache.programs, body, knowledgeCache);
    const response = buildRecommendationResponse(bestRows, body, knowledgeCache);

    return res.status(200).json(response);
  } catch (error) {
    return res.status(500).json({
      error: error.message || "Server error"
    });
  }
}