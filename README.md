# Resume Creator

Data-driven DOCX generation for resumes, references sheets, and cover
letters. Each person's content lives in a single `resume-data.json`; the
scripts in this repo turn that into formatted `.docx` files. Reuse the same
scripts across people/roles — no code changes needed for a new person, only
a new data file.

## Layout

```
resume-creator/
  generate_resumes.py
  generate_references.py
  generate_cover_letters.py
  people/                        # gitignored — contains real personal data
    <Person Name>/
      resume-data.json
      *.docx                     # generated output lands here
```

`people/` is gitignored in its entirety since it contains names, emails,
phone numbers, and other personal data. Only the generator scripts and this
README are meant to be tracked in git.

## Running the generators

All three scripts take one argument: the path to a person's folder
(the folder must contain `resume-data.json`). Run with `uv` so the
`python-docx` dependency is available without a manual venv:

```bash
uv run --with python-docx generate_resumes.py "people/<Person Name>"
uv run --with python-docx generate_references.py "people/<Person Name>"
uv run --with python-docx generate_cover_letters.py "people/<Person Name>"
```

Outputs are written back into that same person's folder.

### generate_resumes.py

Produces four resume layouts, doubled to eight when the newer
`skills.categories` schema is used (see below):

- **Basic** — single-column, traditional layout.
- **Basic OnePage** — same layout, aggressively condensed to fit one page
  (combines Leadership & Volunteer, drops older/irrelevant job bullets to
  header lines, uses `compactBullets`/`compactDetails` where provided).
- **Styled** — two-column layout with a narrow accent sidebar.
- **Styled OnePage** — two-column layout, condensed like Basic OnePage.

If `skills.categories` is present, each of the four layouts is generated
twice — once with `skills_format="narrative"` (dense paragraphs) and once
with `"tags"` (short scannable bullet lists) — producing files suffixed
`- Narrative.docx` / `- Tags.docx`.

The **Styled** two-column layout adapts its column assignment to the data:
if `capstoneAndResearch` or `leadership` is present (student-style schema),
Experience stays in the narrow sidebar and the wide main column showcases
Capstone/Research and Leadership. If neither is present (professional
schema where Experience is the centerpiece), Experience runs in the wide
main column instead, since career-level bullets are usually too long for a
narrow sidebar.

### generate_references.py

Produces a single `<Name> - References.docx`, styled to match the resume
header, from the `references[]` array.

### generate_cover_letters.py

Produces a generic, bracket-placeholder `Cover Letter - Template.docx` for
reuse across applications. Any additional functions in this file
(`build_jnj_fse_letter`, etc.) are one-off letters tailored to a specific
job application for a specific person — copy the pattern for a new
application rather than editing an existing tailored letter in place.

## `resume-data.json` format

Top-level fields:

| Field | Type | Notes |
|---|---|---|
| `name` | string | Full name, used in headers and output filenames. |
| `contact` | object | `phone`, `email`, `location`, optional `headline` (shown next to contact info in the Styled layout's tagline, e.g. "Principal DevSecOps Engineer"). |
| `summary` | string, optional | 2-4 sentence professional summary shown near the top. Omit the field (or see `summaryNote` convention below) to leave it out entirely. |
| `footer` | string, optional | Small italic footer text, e.g. "References available upon request". |
| `skills` | object | See below — two supported shapes. |
| `experience` | array | Job history, most-recent-first after sorting (see below). |
| `education` | array | Degrees/schools. |
| `capstoneAndResearch` | array, optional | Student-schema section; omit entirely for professional resumes. |
| `leadership` | array, optional | Volunteer/leadership roles held separately from paid jobs; omit if leadership responsibilities are folded into `experience` bullets instead. |
| `volunteer` | array, optional | Simple `{role, organization, hours}` entries. |
| `references` | array, optional | Consumed by `generate_references.py`, not by the resume generators. |

### `skills` — two supported shapes

**Newer (categories) schema** — supports both narrative and tag-list
rendering and the 8-file generation mode:

```json
"skills": {
  "categories": [
    {
      "name": "Leadership & Management",
      "narrative": "One paragraph describing this skill area in prose.",
      "tags": ["Short tag one", "Short tag two", "..."]
    }
  ],
  "certifications": ["Cert Name — Month Year", "..."]
}
```

**Legacy (flat list) schema** — still supported, renders narrative-only,
4-file generation mode:

```json
"skills": {
  "technical": ["Skill", "Skill", "..."],
  "technicalGrouped": ["Comma, grouped, skills", "..."],
  "certifications": ["Cert Name", "..."]
}
```

### `experience[]` entries

```json
{
  "title": "Job Title",
  "organization": "Company",
  "location": "City, ST",
  "start": "2024",
  "end": "Present",
  "startFull": "Sept. 2024",
  "endFull": "Present",
  "bullets": ["Full bullet shown on non-condensed layouts", "..."],
  "compactBullets": ["Optional: 1-2 condensed bullets for OnePage variants"],
  "forceCondensed": false,
  "engineeringRelevant": true,
  "keepBulletsOnBasic": false
}
```

- `start`/`end` are short year strings used for sorting and the compact
  date range; `end: "Present"` is recognized and always sorts first.
- `startFull`/`endFull` are the human-readable strings actually printed
  (e.g. "Sept. 2024" instead of "2024"); if omitted, `daterange()` falls
  back to `start`/`end`.
- A job is condensed to a header-only line (no bullets) on Basic/Styled
  layouts when any of: `forceCondensed: true` (explicit override — use
  this for hand-tiering how much detail an old or minor role deserves),
  `engineeringRelevant: false`, or (absent both) it ended more than 4
  years ago. `keepBulletsOnBasic: true` overrides the 4-year auto-condense
  (but not `forceCondensed` or `engineeringRelevant: false`) to keep
  bullets visible on the Basic layout.
- `compactBullets`, if present, replaces `bullets` specifically in the
  OnePage (`full_condense`) variants — use it to give a role 1-2 dense
  sentences instead of a full bullet list when full detail doesn't fit.

**Tiering guidance:** when a career spans many years, don't condense by a
flat cutoff alone — scale detail by a mix of recency, how long the role was
held, and how much impact/breadth it had. A longer, higher-impact role
should generally show more than a short, narrower one even if the short one
is more recent.

### `education[]` / `capstoneAndResearch[]` entries

Same `start`/`end`/`startFull`/`endFull` date fields as experience, plus:

- `education[]`: `institution`, `location`, `details` (list of strings),
  optional `compactDetails` for OnePage variants.
- `capstoneAndResearch[]`: `title`, optional `subtitle`, `bullets`,
  optional `compactBullets`.

### `leadership[]` / `volunteer[]` entries

- `leadership[]`: `title`, `organization`, `start`/`end`/`startFull`/
  `endFull`, `bullets`.
- `volunteer[]`: `role`, `organization`, `hours`.

### `references[]` entries (for `generate_references.py`)

```json
{
  "name": "Reference Name",
  "title": "Their Title",
  "organization": "Their Org",
  "relationship": "How they know you",
  "email": "optional",
  "phone": "optional"
}
```
