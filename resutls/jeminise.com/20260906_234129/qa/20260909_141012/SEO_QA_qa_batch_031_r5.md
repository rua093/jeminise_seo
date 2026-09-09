# SEO QA qa_batch_031_r5

- Source workbook: `resutls\jeminise.com\20260906_234129\revisions\qa_batch_031_r5\SEO_Product_Optimization_qa_batch_031_r5.xlsx`
- SHA-256: `025A68864D79A08435B481EEF8890F82E3EB792CBFAD2543233F866B7F7AB900`
- Admin export: `products_export_1.csv`
- Admin export SHA-256: `729CBD9B3E0248F5D9196D999A49222DBDACA85C2F3CF8524BEEC549B57D738C`
- QA run: `20260909_141012`
- Checked at: `2026-09-09T13:54:42+07:00`
- Scope: 10 products, 71 image positions
- Batch status: `QA_PASS`
- Batch score: `95.2`

## Score Logic
- Chấm theo `seo-prompt/jeminise/prompt_qa.md` v1.0; có mở live storefront/`.js`, đối chiếu admin export, workbook r5 và contact sheet/ảnh local.
- `310`: lỗi `comforter` của QA r4 đã được sửa trong proposed fields, keyword fields và 8 dòng `Image_Audit`; current URL/title vẫn giữ nguyên như dữ liệu nguồn.
- `K3` bị `PARTIAL` cho toàn bộ lô vì bằng chứng nhu cầu mới ở mức `SERP_ONLY`/comparable public SERP.
- `E1` `FULL` cho sản phẩm `310` đã sửa ở r5; `E1 PARTIAL` cho 9 sản phẩm còn lại vì row-level `revision` vẫn ghi `r2` trong workbook r5.
- Không phát hiện `CRITICAL` hoặc `MAJOR`.

| Product | Score | Status | Images | C/M/m | Issue refs |
|---|---:|---|---:|---:|---|
| `personalized-vibrant-cactus-flower-quilt-set-machine-washable-style` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_001; ISSUE_002 |
| `personalized-vibrant-embroidered-desert-cactus-quilt-set-style-8` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_003; ISSUE_004 |
| `personalized-vibrant-flowering-cactus-desert-scene-quilt-set-style` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_005; ISSUE_006 |
| `personalized-vibrant-potted-cactus-garden-quilt-set-machine-washable` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_007; ISSUE_008 |
| `personalized-vibrant-potted-cactus-quilt-set-machine-washa-design-3` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_009; ISSUE_010 |
| `personalized-viking-celtic-tree-of-life-quilt-quilt-set-viking-style` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_011; ISSUE_012 |
| `personalized-vintage-reading-girl-with-flowers-throw-blanket-book` | 95.0 | `QA_PASS` | 6/6 | 0/0/1 | ISSUE_013; ISSUE_014 |
| `personalized-vintage-yggdrasil-tree-of-life-quilt-sets-vintage` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_015; ISSUE_016 |
| `personalized-watercolor-basketball-players-dunking-blanket-83c424fef1-83c424fef1` | 95.0 | `QA_PASS` | 8/8 | 0/0/1 | ISSUE_017; ISSUE_018 |
| `personalized-wildlife-and-romance-two-deer-standing-in-a-forest-comforter-ef860ca9a2-ef860ca9a2` | 97.5 | `QA_PASS` | 8/8 | 0/0/0 | ISSUE_019 |

## Findings
- `meta_description_seo`: không thấy câu bị cắt giữa từ hoặc cắt ý.
- `description_proposed_html`: không thấy câu nội bộ/template như `copy stays specific to the visible artwork`.
- `Image_Audit`: 71/71 vị trí ảnh khớp workbook, admin export và live JSON.
- 10/10 sản phẩm đạt `QA_PASS`; batch `031` hiện không còn lỗi chặn theo rubric.

## Limitations
- Chưa có Search Console, internal search, paid keyword volume hoặc verified review corpus; `K3` không được nâng lên `DEMAND_SUPPORTED`.
- QA không sửa workbook nguồn, không đổi `APPROVED`, không xác nhận import an toàn.
