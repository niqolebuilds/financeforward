# Finance Forward

An internal newsletter for the **Siloam Hospitals Finance & Revenue-Cycle (RCM) team**, teaching agentic AI in a hopeful, human voice. Throughline: *a chatbot answers; an agent acts — and a person still owns the decision.*

This repo is a static site. `index.html` shows the **flagship** version of each issue, grouped into two reading tracks. Alternate **Story-led (A)** and **Data/ROI-led (B)** versions live in `/archive`.

## Structure
- `index.html` — landing page: flagship issues, classified into tracks
  - **Foundations** (Issues 1–3): what agents are, why finance leads, agent anatomy
  - **Proof at Home** (Issues 4–6): revenue cycle, month-end close, treasury/fraud/compliance
- `issues/issue-N.html` — flagship (recommended) article per issue
- `archive/` — alternate A and B versions of every issue
- `assets/style.css` — theme
- `assets/img/*.svg` — original, web-optimized charts (built with matplotlib)

## Status
Issues 1–6 published. Issues 7–15 in progress (the arc continues through change-leadership and an optimistic finale).

## Publishing (GitHub Pages)
Serve the repo root; Pages will pick up `index.html`. `.nojekyll` is included so all asset paths serve as-is.

*Educational internal content. Statistics are cited per issue; sources listed at the foot of each article.*
