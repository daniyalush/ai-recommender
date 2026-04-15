import fs from "fs";
import path from "path";
import xlsx from "xlsx";
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
  if (value === null || value === undefined) return null;
  const cleaned = String(value).replace(/[^0-9.]/g, "");
  const num = parseFloat(cleaned);
  return Number.isFinite(num) ? num : null;
}

function getFirst(row, keys) {
  for (const key of keys) {
    if (row[key] !== undefined && row[key] !== null && String(row[key]).trim() !== "") {
      return row[key];
    }
  }
  return "";
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

function mapProgramRows(rawRows) {
  return rawRows
    .map((row) => {
      const university = getFirst(row, [
        "University", "university", "UNIVERSITY", "Institution", "institution"
      ]);

      const program = getFirst(row, [
  "Program_Title",
  "Program",
  "PROGRAM",
  "program",
  "Programme",
  "PROGRAMME",
  "Major",
  "major",
  "Department",
  "department",
  "Faculty / Program",
  "Faculty/Program",
  "Program Name",
  "PROGRAM NAME",
  "Undergraduate Program",
  "Programs",
  "PROGRAMS",
  "Name",
  "NAME",
  "Title",
  "TITLE"
]);

      const language = getFirst(row, [
        "Language", "language", "Medium", "Instruction Language"
      ]);

      const tuition = getFirst(row, [
        "Tuition", "tuition", "Tuition Fee", "TUITION FEE / PER YEAR",
        "Fee", "Annual Fee", "Price", "price"
      ]);

      const scholarship = getFirst(row, [
        "Scholarship", "scholarship", "Scholarship %", "Scholarship Percentage", "Discount"
      ]);

      const rankingBand = getFirst(row, [
        "Ranking", "ranking", "THE Ranking", "Ranking Band"
      ]);

      if (!university && !program) return null;

      return {
        university: normalizeText(university),
        program: normalizeText(program),
        language: normalizeText(language),
        tuition: normalizeText(tuition),
        scholarship: normalizeText(scholarship),
        ranking_band: normalizeText(rankingBand),
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
    },
    rules: {
      country_scope: "Turkey only",
      intake: "Fall only",
      sat_policy: "SAT is not mandatory for most Turkish universities but may help with scholarships",
      deposit: "Seat reservation deposit is usually around $1,000",
      living_costs: {
        dorm_total_year: "$4,400 to $8,000",
        shared_apartment_total_year: "$6,400 to $10,000",
        food_transport_month: "$300 to $500"
      }
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
  if (pref.includes("campus")) return knowledge.rules.living_costs.dorm_total_year;
  if (pref.includes("city")) return knowledge.rules.living_costs.shared_apartment_total_year;
  return "$4,400 to $10,000";
}

function preferenceScore(row, student) {
  let score = 0;

  const uniPref = safeLower(student.university_preference);
  const intended = safeLower(student.intended_program);
  const uni = safeLower(row.university);
  const program = safeLower(row.program);

  if (uniPref && uni.includes(uniPref)) score += 5;
  if (intended && program.includes(intended)) score += 6;

  const aliases = {
    "computer science": ["computer engineering", "software engineering", "artificial intelligence", "computer programming"],
    "software engineering": ["software engineering", "computer engineering", "back-end software development"],
    "artificial intelligence": ["artificial intelligence", "robotics", "computer engineering"],
    "business": ["business administration", "management", "economics", "finance", "international trade"],
    "psychology": ["psychology", "applied psychology", "clinical psychology"],
    "medicine": ["medicine"],
    "dentistry": ["dentistry"],
    "nursing": ["nursing"],
    "architecture": ["architecture", "interior architecture"],
    "media": ["media", "communication", "journalism", "radio", "television", "visual communication"]
  };

  Object.entries(aliases).forEach(([key, vals]) => {
    if (intended.includes(key) && vals.some((v) => program.includes(v))) {
      score += 3;
    }
  });

  return score;
}

function rankingScore(university, rankingImportance, knowledge) {
  if (safeLower(rankingImportance) !== "yes") return 0;
  return getCategoryPriority(categorizeUniversity(university, knowledge));
}

function scholarshipScore(row, scholarshipPreference) {
  if (!safeLower(scholarshipPreference).includes("yes")) return 0;
  const combined = safeLower(`${row.scholarship} ${row.tuition} ${row.university} ${row.program}`);
  if (combined.includes("75") || combined.includes("60") || combined.includes("50") || combined.includes("scholar")) {
    return 2;
  }
  return 0;
}

function lifestyleScore(university, stylePref) {
  const pref = safeLower(stylePref);
  const uni = safeLower(university);

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

function budgetFitScore(tuitionValue, budgetValue) {
  if (!budgetValue || !tuitionValue) return 0;
  if (tuitionValue <= budgetValue * 0.65) return 3;
  if (tuitionValue <= budgetValue * 0.85) return 2;
  if (tuitionValue <= budgetValue) return 1;
  return -2;
}

function inferFromDocs(docText) {
  const text = safeLower(docText);

  const inference = {
    intended_program: "",
    education_system: "",
    grades_profile: "",
    english_test: "",
    english_score: "",
    sat_score: ""
  };

  if (text.includes("a level") || text.includes("as level")) inference.education_system = "A Levels";
  if (text.includes("o level")) inference.education_system = inference.education_system || "O Levels";
  if (text.includes("fsc")) inference.education_system = "FSC";
  if (text.includes("ib")) inference.education_system = "IB";

  if (text.includes("computer engineering")) inference.intended_program = "Computer Engineering";
  else if (text.includes("software engineering")) inference.intended_program = "Software Engineering";
  else if (text.includes("computer science")) inference.intended_program = "Computer Science";
  else if (text.includes("artificial intelligence")) inference.intended_program = "Artificial Intelligence";
  else if (text.includes("business administration")) inference.intended_program = "Business";
  else if (text.includes("psychology")) inference.intended_program = "Psychology";
  else if (text.includes("medicine")) inference.intended_program = "Medicine";
  else if (text.includes("dentistry")) inference.intended_program = "Dentistry";

  const ieltsMatch = docText.match(/IELTS[^0-9]*([0-9](?:\.[0-9])?)/i);
  if (ieltsMatch) {
    inference.english_test = "IELTS";
    inference.english_score = ieltsMatch[1];
  }

  const satMatch = docText.match(/SAT[^0-9]*([0-9]{3,4})/i);
  if (satMatch) {
    inference.sat_score = satMatch[1];
  }

  inference.grades_profile = docText.slice(0, 800);

  return inference;
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

function buildQuestionResponse(body, missing, docsUploaded) {
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

function scorePrograms(programs, student, knowledge) {
  const budget = toNumber(student.yearly_budget_total);

  const scored = programs.map((row) => {
    const tuitionValue = toNumber(row.tuition);
    const category = categorizeUniversity(row.university, knowledge);

    let score = 0;
    score += preferenceScore(row, student);
    score += rankingScore(row.university, student.ranking_importance, knowledge);
    score += scholarshipScore(row, student.scholarship_preference);
    score += lifestyleScore(row.university, student.lifestyle_preference);
    score += budgetFitScore(tuitionValue, budget);
    score += getCategoryPriority(category);

    return {
      ...row,
      _score: score,
      _category: category
    };
  });

  scored.sort((a, b) => b._score - a._score);

  const seen = new Set();
  const finalRows = [];

  for (const row of scored) {
    const key = `${safeLower(row.university)}|${safeLower(row.program)}`;
    if (!seen.has(key)) {
      seen.add(key);
      finalRows.push(row);
    }
  }

  return finalRows.slice(0, 3);
}

function buildRecommendationResponse(bestRows, student, knowledge, docsUsed) {
  const living = estimateLiving(student.lifestyle_preference, knowledge);

  return {
    mode: "recommendations",
    country_scope: "Turkey only",
    docs_used: docsUsed,
    summary: {
      budget_comment: student.yearly_budget_total
        ? `Recommendations were aligned around an approximate yearly tuition + living budget of ${student.yearly_budget_total}.`
        : "Budget was inferred or not fully specified.",
      ranking_comment: safeLower(student.ranking_importance) === "yes"
        ? "Ranking-sensitive options were prioritized first."
        : "Ranking was not treated as the primary driver.",
      scholarship_comment: "Scholarships are positioned as estimated possibilities, not guarantees.",
      lifestyle_comment: student.lifestyle_preference
        ? `Lifestyle preference considered: ${student.lifestyle_preference}.`
        : "Lifestyle preference was inferred or not fully specified.",
      intake_comment: "For undergraduate Turkey admissions, the relevant intake is Fall only."
    },
    recommendations: bestRows.map((row) => ({
      university: row.university,
      program: row.program,
      category: row._category,
      ranking_band: knowledge.ranking_overrides[row.university] || row.ranking_band || "Indicative only",
      tuition_estimate: row.tuition || "Ask counselor",
      living_cost_estimate: living,
      estimated_total_yearly_cost: "Estimate depends on tuition + living style",
      scholarship_positioning: row.scholarship
        ? `Possible scholarship/discount indicator from source sheet: ${row.scholarship}`
        : "Possible depending on profile and university policy.",
      why_this_fits: [
        `Matches or is close to the student's intended program.`,
        `Fits the Turkey counseling category: ${row._category}.`,
        `Selected using budget, lifestyle, scholarship, and ranking preference logic.`
      ],
      notes: [
        "Undergraduate Turkey intake is Fall only.",
        "Scholarships should be treated as estimated possibilities, not guaranteed."
      ]
    }))
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
    const contentType = req.headers["content-type"] || "";
    let body = {};
    let files = [];

    if (contentType.includes("multipart/form-data")) {
      const parsed = await parseMultipart(req);
      body = parsed.fields || {};
      files = parsed.files || [];
    } else {
      await new Promise((resolve, reject) => {
        let raw = "";
        req.on("data", (chunk) => {
          raw += chunk;
        });
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
      intended_program: body.intended_program || inferred.intended_program || body.course_interest || "",
      grades_profile: body.grades_profile || inferred.grades_profile || "",
      sat_score: body.sat_score || inferred.sat_score || "",
      english_test: body.english_test || inferred.english_test || "",
      english_score: body.english_score || inferred.english_score || ""
    };

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

    const docsOnlyMode = files.length > 0 && Object.values(body).every((v) => !String(v || "").trim());

    if (docsOnlyMode) {
      const docsMissing = [
        { key: "ranking_importance", label: "Is university ranking an important factor for you?" },
        { key: "yearly_budget_total", label: "What is your approximate yearly budget for tuition and living?" },
        { key: "scholarship_preference", label: "Are you specifically looking for high scholarships?" },
        { key: "lifestyle_preference", label: "Do you prefer a city-center university experience or a campus-based environment?" },
        { key: "candidate_type", label: "Are you a private candidate or applying through a school?" }
      ].filter((f) => !String(merged[f.key] || "").trim());

      if (docsMissing.length > 0) {
        return res.status(200).json(buildQuestionResponse(merged, docsMissing, true));
      }
    } else {
      const missing = profilingFields.filter(
        (f) => !String(merged[f.key] || "").trim()
      );

      if (missing.length > 0) {
        return res.status(200).json(buildQuestionResponse(merged, missing, files.length > 0));
      }
    }

    const bestRows = scorePrograms(knowledgeCache.programs, merged, knowledgeCache);
    const response = buildRecommendationResponse(bestRows, merged, knowledgeCache, files.length > 0);

    return res.status(200).json(response);
  } catch (error) {
    return res.status(500).json({
      error: error.message || "Server error"
    });
  }
}