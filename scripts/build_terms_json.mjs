import fs from "node:fs";

const sourcePath = "docs/terms.md";
const outputPath = "data/terms.json";

const FIELD_MAP = {
  "分类": "category",
  "词汇": "term",
  "标准发音": "pronunciation",
  "中文谐音": "chinese_homophone",
  "拆词/来源": "origin",
  "一句话理解": "meaning",
  "中文工作表达": "chinese_work_expression",
  "画面联想": "visual_metaphor",
  "故事记忆": "story",
  "不要误解为": "misconception",
  "相关词区别": "related_terms",
  "常见表达": "common_phrases",
  "真实沟通例句": "examples",
  "图片生成提示词": "image_prompt",
};

const md = fs.readFileSync(sourcePath, "utf8");
const sectionRegex = /^## (\d+)\. (.+)$/gm;
const matches = [...md.matchAll(sectionRegex)];

function clean(value) {
  return value
    .replace(/^```text\s*/g, "")
    .replace(/```\s*$/g, "")
    .replace(/\s+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function parseBracketFields(body) {
  const fields = {};
  const regex = /^【([^】]+)】\n([\s\S]*?)(?=^【[^】]+】\n|^```|\z)/gm;
  for (const match of body.matchAll(regex)) {
    const key = FIELD_MAP[match[1]];
    if (key) fields[key] = clean(match[2]);
  }
  return fields;
}

function parseBulletFields(body) {
  const fields = {};
  const lines = body.split("\n");

  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    const match = line.match(/^- ([^：]+)：(.*)$/);
    if (!match) continue;

    const key = FIELD_MAP[match[1].trim()];
    if (!key) continue;

    const continuation = [];
    while (i + 1 < lines.length && /^  \S/.test(lines[i + 1])) {
      continuation.push(lines[i + 1].trim());
      i += 1;
    }

    const value = [match[2].trim(), ...continuation].filter(Boolean).join("\n");
    fields[key] = clean(value);
  }

  return fields;
}

const terms = matches.map((match, index) => {
  const start = match.index + match[0].length;
  const end = index + 1 < matches.length ? matches[index + 1].index : md.indexOf("\n## 参考来源", start);
  const body = md.slice(start, end === -1 ? undefined : end).trim();
  const imageMatch = body.match(/!\[[^\]]*]\(([^)]+)\)/);
  const image = imageMatch ? imageMatch[1].replace(/^\.\.\//, "") : "";
  const fields = body.includes("【分类】") ? parseBracketFields(body) : parseBulletFields(body);

  return {
    id: Number(match[1]),
    title: match[2].trim(),
    term: fields.term || match[2].trim(),
    category: fields.category || "",
    pronunciation: fields.pronunciation || "",
    chinese_homophone: fields.chinese_homophone || "",
    origin: fields.origin || "",
    meaning: fields.meaning || "",
    chinese_work_expression: fields.chinese_work_expression || "",
    visual_metaphor: fields.visual_metaphor || "",
    story: fields.story || "",
    misconception: fields.misconception || "",
    related_terms: fields.related_terms || "",
    common_phrases: fields.common_phrases || "",
    examples: fields.examples || "",
    image,
    image_prompt: fields.image_prompt || "",
  };
});

fs.writeFileSync(outputPath, `${JSON.stringify(terms, null, 2)}\n`);
console.log(`Wrote ${terms.length} terms to ${outputPath}`);
