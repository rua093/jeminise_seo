# SEO QA qa_batch_029_r5

- Source workbook: `resutls\jeminise.com\20260906_234129\revisions\qa_batch_029_r5\SEO_Product_Optimization_qa_batch_029_r5.xlsx`
- SHA-256: `7A4A0C7DC075616630426BD02847EC01E1026C2B16318C12ED53D880877645F4`
- Admin export: `products_export_1.csv`
- Admin export SHA-256: `729CBD9B3E0248F5D9196D999A49222DBDACA85C2F3CF8524BEEC549B57D738C`
- QA run: `20260909_151928`
- Checked at: `2026-09-09T15:00:40+07:00`
- Scope: 10 products, 74 image positions
- Batch status: `QA_PASS`
- Batch score: `95.2`

## Score Logic
- Chấm theo `seo-prompt/jeminise/prompt_qa.md` v1.0; có mở live storefront/`.js`, đối chiếu admin export, workbook r5 và contact sheet/ảnh local.
- `286`: lỗi `comforter` của QA r4 đã được sửa trong proposed fields, keyword fields và 8 dòng `Image_Audit`; current URL/title vẫn giữ nguyên như dữ liệu nguồn.
- `K3` bị `PARTIAL` cho toàn bộ lô vì bằng chứng nhu cầu mới ở mức `SERP_ONLY`/comparable public SERP.
- `E1` `FULL` cho sản phẩm `286` đã sửa ở r5; `E1 PARTIAL` cho 9 sản phẩm còn lại vì row-level `revision` vẫn ghi `r2` trong workbook r5.
- Không phát hiện `CRITICAL` hoặc `MAJOR`.

| Product | Score | Status | Images | C/M/m | Issue refs |
|---|---:|---|---:|---:|---|
| `personalized-flowering-cactus-succulent-garden-quilt-set-style-7` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_001; ISSUE_002 |
| `personalized-girl-glasses-reading-on-newspaper-pattern-throw-blanket` | 95.0 | `QA_PASS` | 9/9 | 0/0/1 | ISSUE_003; ISSUE_004 |
| `personalized-girl-in-glasses-reading-book-throw-blanket-book-lover` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_005; ISSUE_006 |
| `personalized-girl-messy-bun-reading-book-and-holding-mug-throw-blanket` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_007; ISSUE_008 |
| `personalized-girl-reading-in-front-of-book-bookshelf-throw-design-8` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_009; ISSUE_010 |
| `personalized-hunting-and-outdoor-comforter-with-deer-antlers-4e13e688e0-4e13e688e0` | 97.5 | `QA_PASS` | 8/8 | 0/0/0 | ISSUE_011 |
| `personalized-just-a-girl-who-loves-books-blanket-custom-name-book` | 95.0 | `QA_PASS` | 9/9 | 0/0/1 | ISSUE_012; ISSUE_013 |
| `personalized-love-and-relationships-quilt-with-photo-and-heart-644d558a8f-644d558a8f` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_014; ISSUE_015 |
| `personalized-morning-affirmations-floral-blanket-lightweight` | 95.0 | `QA_PASS` | 6/6 | 0/0/1 | ISSUE_016; ISSUE_017 |
| `personalized-mosaic-tree-of-life-quilt-sets-machine-wa-mosaic-style` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_018; ISSUE_019 |

## Findings
- `meta_description_seo`: không thấy câu bị cắt giữa từ hoặc cắt ý.
- `description_proposed_html`: không thấy câu nội bộ/template như `copy stays specific to the visible artwork`.
- `Image_Audit`: 74/74 vị trí ảnh khớp workbook, admin export và live JSON.
- 10/10 sản phẩm đạt `QA_PASS`; batch `029` hiện không còn lỗi chặn theo rubric.

## Limitations
- Chưa có Search Console, internal search, paid keyword volume hoặc verified review corpus; `K3` không được nâng lên `DEMAND_SUPPORTED`.
- QA không sửa workbook nguồn, không đổi `APPROVED`, không xác nhận import an toàn.
