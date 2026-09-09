# SEO QA qa_batch_033_r4

- Source workbook: `resutls\jeminise.com\20260906_234129\revisions\qa_batch_033_r4\SEO_Product_Optimization_qa_batch_033_r4.xlsx`
- SHA-256: `DD579D05323A57267599887CEF71CF0FE12D507E17A50F4BA799E4CD22634B9B`
- Admin export: `products_export_1.csv`
- Admin export SHA-256: `729CBD9B3E0248F5D9196D999A49222DBDACA85C2F3CF8524BEEC549B57D738C`
- QA run: `20260909_131039`
- Checked at: `2026-09-09T13:10:39+07:00`
- Scope: 10 products, 70 image positions
- Batch status: `QA_PASS`
- Batch score: `93.5`

## Score Logic
- `K3` b? `PARTIAL` cho to?n b? l? v? b?ng ch?ng nhu c?u ch? ? m?c `SERP_ONLY`/comparable public SERP.
- `E1` b? `PARTIAL` cho to?n b? l? v? file ngu?n l? `qa_batch_033_r4` nh?ng c?t `revision` trong t?ng row v?n ghi `r2`.
- `T1` b? `PARTIAL` ri?ng ? c?c s?n ph?m `323`, `327`, `328` do phrasing title h?i k?m t? nhi?n nh?ng kh?ng sai ngh?a.
- Kh?ng ph?t hi?n `CRITICAL` ho?c `MAJOR`.

| Product | Score | Status | Images | Main QA Notes |
|---|---:|---|---:|---|
| `sea-turtle-pattern-quilt-animal-patchwork-sea-turtle-quilt-set` | 95.0 | `QA_PASS` | 7/7 | Proposal correctly replaces misleading live/admin farmhouse snippet with sea turtle patchwork specifics. |
| `softball-ball-sports-comforter-set-bright-teal-yellow-bedding-for` | 95.0 | `QA_PASS` | 7/7 | Proposal separates this item by gray stripes and large yellow softball text rather than generic softball bedding. |
| `softball-bedding-set-for-girls-women-bright-teal-yellow-sports` | 90.0 | `QA_PASS` | 7/7 | Content is truthful, but `Name Number` in SEO title is a bit awkward compared with a smoother personalized softball glove phrasing. |
| `softball-bedding-set-for-softball-lover-ball-sports-theme-bedding-twin` | 95.0 | `QA_PASS` | 7/7 | Strong match to quote patchwork design; no unsupported personalization claim. |
| `softball-bedding-set-for-softball-lover-heart-love-ball-sports-theme` | 95.0 | `QA_PASS` | 7/7 | Strong match to sunflower/player silhouette design; sample name/number is handled as sample artwork. |
| `softball-bedroom-decor-bedding-set-vibrant-teal-yellow-sports-theme` | 95.0 | `QA_PASS` | 7/7 | Good match to pink/yellow quote design with batter silhouettes and splatter graphics. |
| `softball-fan-bedding-set-stylish-teal-yellow-sports-comforter-for` | 90.0 | `QA_PASS` | 7/7 | Visual match is good, but `Name Comforter Set` in title is slightly less natural than `Personalized Vintage Softball Comforter Set`. |
| `softball-lover-comforter-set-teal-yellow-ball-sports-theme-bedding` | 90.0 | `QA_PASS` | 7/7 | Visual match is good, but `Black Yellow` title wording should read more naturally as `Black and Yellow`. |
| `softball-lover-teal-yellow-bedding-stylish-sports-theme-comforter` | 95.0 | `QA_PASS` | 7/7 | Strong match to mint patchwork design; no misleading teal/yellow carryover from current handle/title. |
| `softball-sports-theme-bedding-set-bright-teal-yellow-comforter-for` | 95.0 | `QA_PASS` | 7/7 | Proposal correctly narrows broad current title to yellow stitch design and treats Cayleigh/47 as visible sample details. |

## Findings
- `meta_description_seo`: kh?ng th?y c?u b? c?t gi?a t?/?; to?n b? meta l? c?u ho?n ch?nh v? b?m thi?t k? c? th?.
- `description_proposed_html`: kh?ng th?y c?u n?i b?/template nh? `copy stays specific to the visible artwork`.
- `Image_Audit`: 70/70 v? tr? ?nh kh?p contact sheet, admin export v? live JSON; alt ?? xu?t c? th? h?n alt hi?n t?i trong admin export.
- `batch_033` ??t `QA_PASS` theo rubric v? m?i s?n ph?m tr?n 85 ?i?m v? kh?ng c? `CRITICAL/MAJOR`.

## Issues And Limitations
- `ISSUE_001` `LIMITATION` `sea-turtle-pattern-quilt-animal-patchwork-sea-turtle-quilt-set` `keyword_evidence_level`: Demand fit is plausible and truthful, but cannot be fully validated without Search Console, internal search, paid volume, or verified customer review corpus.
- `ISSUE_002` `MINOR` `sea-turtle-pattern-quilt-animal-patchwork-sea-turtle-quilt-set` `revision`: Revision metadata is inconsistent with the frozen workbook path. It does not change product copy correctness, but weakens traceability.
- `ISSUE_003` `LIMITATION` `softball-ball-sports-comforter-set-bright-teal-yellow-bedding-for` `keyword_evidence_level`: Demand fit is plausible and truthful, but cannot be fully validated without Search Console, internal search, paid volume, or verified customer review corpus.
- `ISSUE_004` `MINOR` `softball-ball-sports-comforter-set-bright-teal-yellow-bedding-for` `revision`: Revision metadata is inconsistent with the frozen workbook path. It does not change product copy correctness, but weakens traceability.
- `ISSUE_005` `LIMITATION` `softball-bedding-set-for-girls-women-bright-teal-yellow-sports` `keyword_evidence_level`: Demand fit is plausible and truthful, but cannot be fully validated without Search Console, internal search, paid volume, or verified customer review corpus.
- `ISSUE_006` `MINOR` `softball-bedding-set-for-girls-women-bright-teal-yellow-sports` `revision`: Revision metadata is inconsistent with the frozen workbook path. It does not change product copy correctness, but weakens traceability.
- `ISSUE_007` `MINOR` `softball-bedding-set-for-girls-women-bright-teal-yellow-sports` `meta_title_seo`: `Name Number` is understandable but unnatural; better phrasing would be `Personalized Softball Glove Comforter Set`.
- `ISSUE_008` `LIMITATION` `softball-bedding-set-for-softball-lover-ball-sports-theme-bedding-twin` `keyword_evidence_level`: Demand fit is plausible and truthful, but cannot be fully validated without Search Console, internal search, paid volume, or verified customer review corpus.
- `ISSUE_009` `MINOR` `softball-bedding-set-for-softball-lover-ball-sports-theme-bedding-twin` `revision`: Revision metadata is inconsistent with the frozen workbook path. It does not change product copy correctness, but weakens traceability.
- `ISSUE_010` `LIMITATION` `softball-bedding-set-for-softball-lover-heart-love-ball-sports-theme` `keyword_evidence_level`: Demand fit is plausible and truthful, but cannot be fully validated without Search Console, internal search, paid volume, or verified customer review corpus.
- `ISSUE_011` `MINOR` `softball-bedding-set-for-softball-lover-heart-love-ball-sports-theme` `revision`: Revision metadata is inconsistent with the frozen workbook path. It does not change product copy correctness, but weakens traceability.
- `ISSUE_012` `LIMITATION` `softball-bedroom-decor-bedding-set-vibrant-teal-yellow-sports-theme` `keyword_evidence_level`: Demand fit is plausible and truthful, but cannot be fully validated without Search Console, internal search, paid volume, or verified customer review corpus.
- `ISSUE_013` `MINOR` `softball-bedroom-decor-bedding-set-vibrant-teal-yellow-sports-theme` `revision`: Revision metadata is inconsistent with the frozen workbook path. It does not change product copy correctness, but weakens traceability.
- `ISSUE_014` `LIMITATION` `softball-fan-bedding-set-stylish-teal-yellow-sports-comforter-for` `keyword_evidence_level`: Demand fit is plausible and truthful, but cannot be fully validated without Search Console, internal search, paid volume, or verified customer review corpus.
- `ISSUE_015` `MINOR` `softball-fan-bedding-set-stylish-teal-yellow-sports-comforter-for` `revision`: Revision metadata is inconsistent with the frozen workbook path. It does not change product copy correctness, but weakens traceability.
- `ISSUE_016` `MINOR` `softball-fan-bedding-set-stylish-teal-yellow-sports-comforter-for` `meta_title_seo`: `Name Comforter Set` is less natural than `Personalized Vintage Softball Comforter Set`.
- `ISSUE_017` `LIMITATION` `softball-lover-comforter-set-teal-yellow-ball-sports-theme-bedding` `keyword_evidence_level`: Demand fit is plausible and truthful, but cannot be fully validated without Search Console, internal search, paid volume, or verified customer review corpus.
- `ISSUE_018` `MINOR` `softball-lover-comforter-set-teal-yellow-ball-sports-theme-bedding` `revision`: Revision metadata is inconsistent with the frozen workbook path. It does not change product copy correctness, but weakens traceability.
- `ISSUE_019` `MINOR` `softball-lover-comforter-set-teal-yellow-ball-sports-theme-bedding` `meta_title_seo`: `Black Yellow` is missing `and`; better `Black and Yellow Softball Flag Comforter Set`.
- `ISSUE_020` `LIMITATION` `softball-lover-teal-yellow-bedding-stylish-sports-theme-comforter` `keyword_evidence_level`: Demand fit is plausible and truthful, but cannot be fully validated without Search Console, internal search, paid volume, or verified customer review corpus.
- `ISSUE_021` `MINOR` `softball-lover-teal-yellow-bedding-stylish-sports-theme-comforter` `revision`: Revision metadata is inconsistent with the frozen workbook path. It does not change product copy correctness, but weakens traceability.
- `ISSUE_022` `LIMITATION` `softball-sports-theme-bedding-set-bright-teal-yellow-comforter-for` `keyword_evidence_level`: Demand fit is plausible and truthful, but cannot be fully validated without Search Console, internal search, paid volume, or verified customer review corpus.
- `ISSUE_023` `MINOR` `softball-sports-theme-bedding-set-bright-teal-yellow-comforter-for` `revision`: Revision metadata is inconsistent with the frozen workbook path. It does not change product copy correctness, but weakens traceability.