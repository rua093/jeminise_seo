# SEO QA qa_batch_029_r4

- Source workbook: `resutls\jeminise.com\20260906_234129\revisions\qa_batch_029_r4\SEO_Product_Optimization_qa_batch_029_r4.xlsx`
- SHA-256: `3647AB68AA5F26C6FDC28A73A77FBE2BDCA67269C9DBAA6A0012A8DD2EEB6ABE`
- Admin export: `products_export_1.csv`
- Admin export SHA-256: `729CBD9B3E0248F5D9196D999A49222DBDACA85C2F3CF8524BEEC549B57D738C`
- QA run: `20260909_145706`
- Checked at: `2026-09-09T14:49:35+07:00`
- Scope: 10 products, 74 image positions
- Batch status: `QA_REVISE`
- Batch score: `92.8`

## Score Logic
- Chấm theo `seo-prompt/jeminise/prompt_qa.md` v1.0; có mở live storefront/`.js`, đối chiếu admin export, workbook r4 và contact sheet/ảnh local.
- `286` bị `MAJOR`: proposed fields, keyword map và 8 alt/observation dùng `comforter/comforter set`, trong khi admin `Type`, live JSON type, option size, live description và contact-sheet info panels support `Quilt/Quilt Set`.
- `K3` bị `PARTIAL` cho toàn bộ lô vì bằng chứng nhu cầu mới ở mức `SERP_ONLY`/comparable public SERP.
- `E1` bị `PARTIAL` cho toàn bộ lô vì source workbook là `qa_batch_029_r4` nhưng row-level `revision` vẫn ghi `r2`.
- Không phát hiện `CRITICAL`.

| Product | Score | Status | Images | C/M/m | Issue refs |
|---|---:|---|---:|---:|---|
| `personalized-flowering-cactus-succulent-garden-quilt-set-style-7` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_001; ISSUE_002 |
| `personalized-girl-glasses-reading-on-newspaper-pattern-throw-blanket` | 95.0 | `QA_PASS` | 9/9 | 0/0/1 | ISSUE_003; ISSUE_004 |
| `personalized-girl-in-glasses-reading-book-throw-blanket-book-lover` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_005; ISSUE_006 |
| `personalized-girl-messy-bun-reading-book-and-holding-mug-throw-blanket` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_007; ISSUE_008 |
| `personalized-girl-reading-in-front-of-book-bookshelf-throw-design-8` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_009; ISSUE_010 |
| `personalized-hunting-and-outdoor-comforter-with-deer-antlers-4e13e688e0-4e13e688e0` | 73.0 | `QA_REVISE` | 8/8 | 0/1/1 | ISSUE_011; ISSUE_012; ISSUE_013 |
| `personalized-just-a-girl-who-loves-books-blanket-custom-name-book` | 95.0 | `QA_PASS` | 9/9 | 0/0/1 | ISSUE_014; ISSUE_015 |
| `personalized-love-and-relationships-quilt-with-photo-and-heart-644d558a8f-644d558a8f` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_016; ISSUE_017 |
| `personalized-morning-affirmations-floral-blanket-lightweight` | 95.0 | `QA_PASS` | 6/6 | 0/0/1 | ISSUE_018; ISSUE_019 |
| `personalized-mosaic-tree-of-life-quilt-sets-machine-wa-mosaic-style` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_020; ISSUE_021 |

## Findings
- `meta_description_seo`: không thấy câu bị cắt giữa từ hoặc cắt ý.
- `description_proposed_html`: không thấy câu nội bộ/template như `copy stays specific to the visible artwork`.
- `Image_Audit`: 74/74 vị trí ảnh khớp workbook, admin export và live JSON.
- 9 sản phẩm `281-285` và `287-290` đạt `QA_PASS`; sản phẩm `286` cần revision trước khi batch có thể pass.

## Priority Fix
1. Sửa sản phẩm `286` để dùng `Quilt`/`Quilt Set` nhất quán trong `title_proposed`, `meta_title_seo`, `meta_description_seo`, `description_proposed_html`, keyword fields, `Keyword_Map`, `Product_Evidence`, `Buyer_Search_Research` và 8 dòng `Image_Audit`.
2. Cập nhật metadata `revision` trong row lên đúng revision mới khi tạo bản sửa tiếp theo.

## Limitations
- Chưa có Search Console, internal search, paid keyword volume hoặc verified review corpus; `K3` không được nâng lên `DEMAND_SUPPORTED`.
- QA không sửa workbook nguồn, không đổi `APPROVED`, không xác nhận import an toàn.
