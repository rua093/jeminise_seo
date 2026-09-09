# SEO QA qa_batch_030_r4

- Source workbook: `resutls\jeminise.com\20260906_234129\revisions\qa_batch_030_r4\SEO_Product_Optimization_qa_batch_030_r4.xlsx`
- SHA-256: `19058334B8B7421D3D59EBC00239F1F37B0D4F77E37C5665A7CF52702E854E7F`
- Admin export: `products_export_1.csv`
- Admin export SHA-256: `729CBD9B3E0248F5D9196D999A49222DBDACA85C2F3CF8524BEEC549B57D738C`
- QA run: `20260909_142635`
- Checked at: `2026-09-09T14:02:15+07:00`
- Scope: 10 products, 70 image positions
- Batch status: `QA_REVISE`
- Batch score: `92.8`

## Score Logic
- Chấm theo `seo-prompt/jeminise/prompt_qa.md` v1.0; có mở live storefront/`.js`, đối chiếu admin export, workbook r4 và contact sheet/ảnh local.
- `295` bị `MAJOR`: proposed fields, keyword map và 5 alt/observation dùng `comforter/comforter set`, trong khi admin `Type`, live JSON type, option size, live description và contact-sheet info panels support `Quilt/Quilt Set`.
- `K3` bị `PARTIAL` cho toàn bộ lô vì bằng chứng nhu cầu mới ở mức `SERP_ONLY`/comparable public SERP.
- `E1` bị `PARTIAL` cho toàn bộ lô vì source workbook là `qa_batch_030_r4` nhưng row-level `revision` vẫn ghi `r2`.
- Không phát hiện `CRITICAL`.

| Product | Score | Status | Images | C/M/m | Issue refs |
|---|---:|---|---:|---:|---|
| `personalized-reading-girl-on-book-with-flowers-blanket-book-lover` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_001; ISSUE_002 |
| `personalized-reading-girl-on-heart-bookshelf-throw-blanket-book` | 95.0 | `QA_PASS` | 9/9 | 0/0/1 | ISSUE_003; ISSUE_004 |
| `personalized-s-day-mother-and-daughter-quilt-with-photo-and-daisies-8dd40d7242-8dd40d7242` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_005; ISSUE_006 |
| `personalized-seasonal-bookworms-sophia-reading-girl-throw-blanket` | 95.0 | `QA_PASS` | 9/9 | 0/0/1 | ISSUE_007; ISSUE_008 |
| `personalized-southwestern-wolf-head-with-geometric-patterns-comforter-1d03ff91e5-1d03ff91e5` | 73.4 | `QA_REVISE` | 5/5 | 0/1/1 | ISSUE_009; ISSUE_010; ISSUE_011 |
| `personalized-sunflower-christian-blanket-with-bible-verses-custom` | 95.0 | `QA_PASS` | 6/6 | 0/0/1 | ISSUE_012; ISSUE_013 |
| `personalized-sunflower-god-says-you-are-throw-blanket-lightweight` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_014; ISSUE_015 |
| `personalized-sunflowers-and-butterflies-affirmations-blanket-unique` | 95.0 | `QA_PASS` | 6/6 | 0/0/1 | ISSUE_016; ISSUE_017 |
| `personalized-teal-argyle-reading-girl-with-books-throw-blanket-book` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_018; ISSUE_019 |
| `personalized-vibrant-cactus-and-desert-flowers-quilt-set-style-4` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_020; ISSUE_021 |

## Findings
- `meta_description_seo`: không thấy câu bị cắt giữa từ hoặc cắt ý.
- `description_proposed_html`: không thấy câu nội bộ/template như `copy stays specific to the visible artwork`.
- `Image_Audit`: 70/70 vị trí ảnh khớp workbook, admin export và live JSON.
- 9 sản phẩm `291-294` và `296-300` đạt `QA_PASS`; sản phẩm `295` cần revision trước khi batch có thể pass.

## Priority Fix
1. Sửa sản phẩm `295` để dùng `Quilt`/`Quilt Set` nhất quán trong `title_proposed`, `meta_title_seo`, `meta_description_seo`, `description_proposed_html`, keyword fields, `Keyword_Map`, `Product_Evidence`, `Buyer_Search_Research` và 5 dòng `Image_Audit`.
2. Cập nhật metadata `revision` trong row lên đúng revision mới khi tạo bản sửa tiếp theo.

## Limitations
- Chưa có Search Console, internal search, paid keyword volume hoặc verified review corpus; `K3` không được nâng lên `DEMAND_SUPPORTED`.
- QA không sửa workbook nguồn, không đổi `APPROVED`, không xác nhận import an toàn.
