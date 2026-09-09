# SEO QA qa_batch_030_r5

- Source workbook: `resutls\jeminise.com\20260906_234129\revisions\qa_batch_030_r5\SEO_Product_Optimization_qa_batch_030_r5.xlsx`
- SHA-256: `C18187523EDF34761822DCCBA9A663BDF5EFEBF456A8E8BFB481D82D0F2DC76C`
- Admin export: `products_export_1.csv`
- Admin export SHA-256: `729CBD9B3E0248F5D9196D999A49222DBDACA85C2F3CF8524BEEC549B57D738C`
- QA run: `20260909_143812`
- Checked at: `2026-09-09T14:40:02+07:00`
- Scope: 10 products, 70 image positions
- Batch status: `QA_PASS`
- Batch score: `95.2`

## Score Logic
- Chấm theo `seo-prompt/jeminise/prompt_qa.md` v1.0; có mở live storefront/`.js`, đối chiếu admin export, workbook r5 và contact sheet/ảnh local.
- `295`: lỗi `comforter` của QA r4 đã được sửa trong proposed fields, keyword fields và 5 dòng `Image_Audit`; current URL vẫn giữ nguyên như dữ liệu nguồn.
- `K3` bị `PARTIAL` cho toàn bộ lô vì bằng chứng nhu cầu mới ở mức `SERP_ONLY`/comparable public SERP.
- `E1` `FULL` cho sản phẩm `295` đã sửa ở r5; `E1 PARTIAL` cho 9 sản phẩm còn lại vì row-level `revision` vẫn ghi `r2` trong workbook r5.
- Không phát hiện `CRITICAL` hoặc `MAJOR`.

| Product | Score | Status | Images | C/M/m | Issue refs |
|---|---:|---|---:|---:|---|
| `personalized-reading-girl-on-book-with-flowers-blanket-book-lover` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_001; ISSUE_002 |
| `personalized-reading-girl-on-heart-bookshelf-throw-blanket-book` | 95.0 | `QA_PASS` | 9/9 | 0/0/1 | ISSUE_003; ISSUE_004 |
| `personalized-s-day-mother-and-daughter-quilt-with-photo-and-daisies-8dd40d7242-8dd40d7242` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_005; ISSUE_006 |
| `personalized-seasonal-bookworms-sophia-reading-girl-throw-blanket` | 95.0 | `QA_PASS` | 9/9 | 0/0/1 | ISSUE_007; ISSUE_008 |
| `personalized-southwestern-wolf-head-with-geometric-patterns-comforter-1d03ff91e5-1d03ff91e5` | 97.5 | `QA_PASS` | 5/5 | 0/0/0 | ISSUE_009 |
| `personalized-sunflower-christian-blanket-with-bible-verses-custom` | 95.0 | `QA_PASS` | 6/6 | 0/0/1 | ISSUE_010; ISSUE_011 |
| `personalized-sunflower-god-says-you-are-throw-blanket-lightweight` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_012; ISSUE_013 |
| `personalized-sunflowers-and-butterflies-affirmations-blanket-unique` | 95.0 | `QA_PASS` | 6/6 | 0/0/1 | ISSUE_014; ISSUE_015 |
| `personalized-teal-argyle-reading-girl-with-books-throw-blanket-book` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_016; ISSUE_017 |
| `personalized-vibrant-cactus-and-desert-flowers-quilt-set-style-4` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_018; ISSUE_019 |

## Findings
- `meta_description_seo`: không thấy câu bị cắt giữa từ hoặc cắt ý.
- `description_proposed_html`: không thấy câu nội bộ/template như `copy stays specific to the visible artwork`.
- `Image_Audit`: 70/70 vị trí ảnh khớp workbook, admin export và live JSON.
- 10/10 sản phẩm đạt `QA_PASS`; batch `030` hiện không còn lỗi chặn theo rubric.

## Limitations
- Chưa có Search Console, internal search, paid keyword volume hoặc verified review corpus; `K3` không được nâng lên `DEMAND_SUPPORTED`.
- QA không sửa workbook nguồn, không đổi `APPROVED`, không xác nhận import an toàn.
