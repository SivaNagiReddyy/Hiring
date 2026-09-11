// Mirrors the `FACTORS` array in data-hiring-guidelines.html (key -> column
// name only, hints/roles don't matter server-side). Keep this in sync if a
// factor is ever added, renamed, or removed in the frontend, since it decides
// the Scorecards sheet's column headers.
module.exports = [
  { key: "comm", name: "Communication & consultative skill" },
  { key: "sql", name: "Hands-on SQL & PySpark" },
  { key: "model", name: "Data modelling & warehousing" },
  { key: "platform", name: "Platform knowledge & breadth" },
  { key: "arch", name: "Architecture & scenario depth" },
  { key: "decision", name: "Technical decision-making & questioning the status quo" },
  { key: "lead", name: "Team leading & mentoring" },
  { key: "currency", name: "Currency in Data & AI (incl. GenAI)" },
  { key: "learning", name: "Certifications & continuous learning" },
  { key: "govcost", name: "Cost, governance & security at scale" }
];
