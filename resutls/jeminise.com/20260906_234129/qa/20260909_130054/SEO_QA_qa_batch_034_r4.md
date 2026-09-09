# SEO QA qa_batch_034_r4 - Deep Rerun

- Source workbook: `resutls\jeminise.com\20260906_234129\revisions\qa_batch_034_r4\SEO_Product_Optimization_qa_batch_034_r4.xlsx`
- SHA-256: `4735EC320645E865905C564710F988006CB2428756154093B75B69599F6FE335`
- Admin export: `products_export_1.csv`
- Admin export SHA-256: `729CBD9B3E0248F5D9196D999A49222DBDACA85C2F3CF8524BEEC549B57D738C`
- QA run: `20260909_130054`
- Checked at: `2026-09-09T13:00:54+07:00`
- Scope: 8 products, 57 image positions
- Batch status: `QA_PASS`
- Batch score: `96.9`

## Score Logic
- Kh?ng d?ng l?i hai b?n QA c?; ?? x?a v? t?o b?n m?i.
- `K3` b? `PARTIAL` cho to?n b? l? v? b?ng ch?ng nhu c?u ch? ? m?c `SERP_ONLY`, kh?ng c? Search Console/internal search/paid volume/review corpus.
- `E1` ???c ch?m `FULL` v? l??t n?y ?? ??i chi?u tr?c ti?p `products_export_1.csv` cho admin SEO/current image fields.
- Kh?ng ph?t hi?n `CRITICAL` ho?c `MAJOR`; s?n ph?m `337` c? m?t `MINOR` ? phrasing SEO title n?n b? tr? th?m ? `T1`.

| Product | Score | Status | Images | Main QA Notes |
|---|---:|---|---:|---|
| `softball-themed-bedding-for-girls-and-women-bright-teal-yellow` | 97.5 | `QA_PASS` | 7/7 | Good recovery from broad current title: proposed keyword/title focuses on the yellow distressed softball artwork rather than generic girls/women bedding. |
| `sports-theme-softball-bedding-premium-teal-yellow-comforter-set-for` | 97.5 | `QA_PASS` | 7/7 | Good separation from nearby softball items by using polka dot, teal stripe and paisley/glove artwork. |
| `teal-yellow-softball-comforter-set-vibrant-sports-bedroom-decor` | 97.5 | `QA_PASS` | 7/7 | Good separation through fireball/flame language; personalization samples are treated as sample artwork, not guaranteed default text. |
| `teal-yellow-softball-bedding-set-stylish-sports-bedroom-decor` | 97.5 | `QA_PASS` | 14/14 | Good design-specific rewrite; duplicate live images are acknowledged and alt remains suitable per repeated position. |
| `tree-of-life-with-celtic-knot-roots-quilt-set-and-celtic-knot-borders-9482d2b2f5-9482d2b2f5` | 97.5 | `QA_PASS` | 7/7 | Good concise title; description preserves Celtic roots/border details and avoids overclaiming handmade/craft quality. |
| `winter-cardinal-on-berry-branch-patchwork-christmas-quilt-43ecda62b6-43ecda62b6` | 97.5 | `QA_PASS` | 5/5 | Good single-cardinal distinction; title/meta avoid claiming custom name as default artwork. |
| `winter-cardinals-on-heart-shaped-berry-branches-patchwork-christmas-quilt-de1c648f60-de1c648f60` | 92.5 | `QA_PASS` | 5/5 | Mostly strong, but SEO title phrase "Heart Branch" is less natural/precise than "Heart-Shaped Berry Branch". Minor: `Heart Branch` is less natural than `heart-shaped berry branches`. |
| `winter-cat-and-cardinal-sitting-on-snowy-patchwork-christmas-quilt-82f86611f4-82f86611f4` | 97.5 | `QA_PASS` | 5/5 | Good inclusion of visible faith-inspired text; does not turn it into unsupported religious/product claim. |

## Findings
- `meta_description_seo`: kh?ng th?y c?u b? c?t gi?a t?/?; to?n b? meta l? c?u ho?n ch?nh v? b?m thi?t k? c? th?.
- `description_proposed_html`: kh?ng th?y c?u n?i b?/template nh? `copy stays specific to the visible artwork`.
- `Image_Audit`: 57/57 v? tr? ?nh kh?p contact sheet v? admin export/live count; alt ?? xu?t c? th? h?n alt hi?n t?i trong admin export.
- `batch_034` ??t `QA_PASS` theo rubric v? m?i s?n ph?m tr?n 85 ?i?m v? kh?ng c? `CRITICAL/MAJOR`.

## Issues And Limitations
- `ISSUE_001` `LIMITATION` `softball-themed-bedding-for-girls-and-women-bright-teal-yellow` `keyword_evidence_level`: No Search Console, internal-search logs, paid volume, or verified review corpus was available, so demand strength cannot be fully validated.
- `ISSUE_002` `LIMITATION` `sports-theme-softball-bedding-premium-teal-yellow-comforter-set-for` `keyword_evidence_level`: No Search Console, internal-search logs, paid volume, or verified review corpus was available, so demand strength cannot be fully validated.
- `ISSUE_003` `LIMITATION` `teal-yellow-softball-comforter-set-vibrant-sports-bedroom-decor` `keyword_evidence_level`: No Search Console, internal-search logs, paid volume, or verified review corpus was available, so demand strength cannot be fully validated.
- `ISSUE_004` `LIMITATION` `teal-yellow-softball-bedding-set-stylish-sports-bedroom-decor` `keyword_evidence_level`: No Search Console, internal-search logs, paid volume, or verified review corpus was available, so demand strength cannot be fully validated.
- `ISSUE_005` `LIMITATION` `tree-of-life-with-celtic-knot-roots-quilt-set-and-celtic-knot-borders-9482d2b2f5-9482d2b2f5` `keyword_evidence_level`: No Search Console, internal-search logs, paid volume, or verified review corpus was available, so demand strength cannot be fully validated.
- `ISSUE_006` `LIMITATION` `winter-cardinal-on-berry-branch-patchwork-christmas-quilt-43ecda62b6-43ecda62b6` `keyword_evidence_level`: No Search Console, internal-search logs, paid volume, or verified review corpus was available, so demand strength cannot be fully validated.
- `ISSUE_007` `LIMITATION` `winter-cardinals-on-heart-shaped-berry-branches-patchwork-christmas-quilt-de1c648f60-de1c648f60` `keyword_evidence_level`: No Search Console, internal-search logs, paid volume, or verified review corpus was available, so demand strength cannot be fully validated.
- `ISSUE_008` `MINOR` `winter-cardinals-on-heart-shaped-berry-branches-patchwork-christmas-quilt-de1c648f60-de1c648f60` `meta_title_seo`: Phrase `Heart Branch` is slightly less natural and less precise than the visible heart-shaped berry branches, though meaning remains understandable.
- `ISSUE_009` `LIMITATION` `winter-cat-and-cardinal-sitting-on-snowy-patchwork-christmas-quilt-82f86611f4-82f86611f4` `keyword_evidence_level`: No Search Console, internal-search logs, paid volume, or verified review corpus was available, so demand strength cannot be fully validated.