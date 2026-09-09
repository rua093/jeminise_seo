# SEO QA qa_batch_031_r4

- Source workbook: `resutls\jeminise.com\20260906_234129\revisions\qa_batch_031_r4\SEO_Product_Optimization_qa_batch_031_r4.xlsx`
- SHA-256: `9ED591B28F4B7BE0893FFF585898EAF8D918C0CD6D326AA71689019AB21B14E8`
- Admin export: `products_export_1.csv`
- Admin export SHA-256: `729CBD9B3E0248F5D9196D999A49222DBDACA85C2F3CF8524BEEC549B57D738C`
- QA run: `20260909_135826`
- Checked at: `2026-09-09T13:45:51+07:00`
- Scope: 10 products, 71 image positions
- Batch status: `QA_REVISE`
- Batch score: `92.8`

## Score Logic
- Chấm theo `seo-prompt/jeminise/prompt_qa.md` v1.0; có mở live storefront/`.js`, đối chiếu admin export, workbook r4 và contact sheet/ảnh local.
- `310` bị `MAJOR` vì proposed fields, keyword map và 8 alt/observation dùng `comforter/comforter set`, trong khi admin `Type`, live JSON type, option size, live/admin description và contact-sheet info panels support `Quilt/Quilt Set`.
- `K3` bị `PARTIAL` cho toàn bộ lô vì bằng chứng nhu cầu mới ở mức `SERP_ONLY`/comparable public SERP.
- `E1` bị `PARTIAL` cho toàn bộ lô vì source workbook là `qa_batch_031_r4` nhưng row-level `revision` vẫn ghi `r2`.
- Không phát hiện `CRITICAL`.

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
| `personalized-wildlife-and-romance-two-deer-standing-in-a-forest-comforter-ef860ca9a2-ef860ca9a2` | 73.0 | `QA_REVISE` | 8/8 | 0/1/1 | ISSUE_019; ISSUE_020; ISSUE_021 |

## Findings
- `meta_description_seo`: không thấy câu bị cắt giữa từ hoặc cắt ý.
- `description_proposed_html`: không thấy câu nội bộ/template như `copy stays specific to the visible artwork`.
- `Image_Audit`: 71/71 vị trí ảnh khớp workbook, admin export và live JSON.
- 9 sản phẩm `301-309` đạt `QA_PASS`; sản phẩm `310` cần revision trước khi batch có thể pass.

## Priority Fix
1. Sửa sản phẩm `310` để dùng `Quilt`/`Quilt Set` nhất quán trong `title_proposed`, `meta_title_seo`, `meta_description_seo`, `description_proposed_html`, keyword fields, `Keyword_Map`, `Product_Evidence`, `Buyer_Search_Research` và 8 dòng `Image_Audit`.
2. Cập nhật metadata `revision` trong row lên đúng revision mới khi tạo bản sửa tiếp theo.

## Limitations
- Chưa có Search Console, internal search, paid keyword volume hoặc verified review corpus; `K3` không được nâng lên `DEMAND_SUPPORTED`.
- QA không sửa workbook nguồn, không đổi `APPROVED`, không xác nhận import an toàn.
