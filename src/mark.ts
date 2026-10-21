/** Names that are a chapter, not a company. */
const SKIP =
  /^(slt|voice|chat|chatbot|iva|containment|authentication|approval|untitled|scorecard|briefing|sheet|figures|results|the book|northline care|harbor cover)$/i;

/** Domains the Product Hub already trusts for Logo.dev. A guess is not used. */
const DOMAIN: Record<string, string> = {
  samsung: "samsung.com",
  walmart: "walmart.com",
  luxottica: "luxottica.com",
  dish: "dish.com",
  "boost mobile": "boostmobile.com",
  boost: "boostmobile.com",
  santander: "santander.com",
  bupa: "bupa.com",
  surebright: "surebright.com",
  newegg: "newegg.com",
  rogers: "rogers.com",
  sasktel: "sasktel.com",
  barclays: "barclays.co.uk",
  amazon: "amazon.com",
  farmers: "farmers.com",
  "john lewis": "johnlewis.com",
  "t-mobile": "t-mobile.com",
  tmobile: "t-mobile.com",
  verizon: "verizon.com",
  asurion: "asurion.com",
  optus: "optus.com.au",
  "freedom mobile": "freedommobile.ca",
  ee: "ee.co.uk",
  o2: "o2.co.uk",
  apple: "apple.com",
  aviva: "aviva.com",
  natwest: "natwest.com",
  zurich: "zurich.com",
  glentel: "glentel.com",
  "deutsche telekom": "telekom.com",
  videotron: "videotron.com",
  "c spire": "cspire.com",
  sprint: "sprint.com",
  loblaw: "loblaw.ca",
};

/** Approved marks shipped with Product Hub. These win over a searched logo. */
const FILE: Record<string, string> = {
  walmart: "walmart.svg",
  samsung: "samsung.svg",
  luxottica: "luxottica-us.svg",
  dish: "dish.svg",
  santander: "santander.svg",
  surebright: "surebright.svg",
  newegg: "newegg.svg",
  rogers: "rogers.svg",
  sasktel: "sasktel.svg",
  barclays: "barclays.svg",
  farmers: "farmers-home-tech.svg",
  "john lewis": "john-lewis.svg",
  optus: "optus.svg",
  "freedom mobile": "freedom-mobile.svg",
  ee: "ee-flex.svg",
  o2: "o2.svg",
  apple: "apple.svg",
  aviva: "aviva.svg",
  natwest: "natwest.svg",
  zurich: "zurich.svg",
  glentel: "glentel.svg",
  "deutsche telekom": "deutsche-telekom.svg",
  videotron: "videotron.svg",
  "c spire": "cspire.png",
  sprint: "sprint.svg",
  loblaw: "loblaw.svg",
  bupa: "bupa-home-tech.svg",
};

export type MarkRequest = {
  name: string;
  domain: string | null;
  file: string | null;
};

function keyOf(name: string): string {
  return name.trim().toLowerCase().replace(/\s+/g, " ");
}

function lookup<T>(name: string, table: Record<string, T>): T | null {
  const key = keyOf(name);
  if (table[key]) return table[key];
  const words = ` ${key.replace(/[^a-z0-9]+/g, " ").trim()} `;
  for (const [known, value] of Object.entries(table)) {
    if (words.includes(` ${known} `)) return value;
  }
  return null;
}

/** A client the deck can look up. Greetings and chapter titles are not clients. */
export function markRequest(name: string): MarkRequest | null {
  const clean = name.trim();
  if (clean.length < 2 || clean.length > 48) return null;
  if (SKIP.test(clean) || /^(hi|hey|hello|dear|can you|could you|please)\b/i.test(clean)) return null;
  const domain = lookup(clean, DOMAIN);
  const file = lookup(clean, FILE);
  if (!domain && !file) return null;
  return { name: clean, domain, file };
}
