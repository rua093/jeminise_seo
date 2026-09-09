# SEO QA qa_batch_032_r5

- Source workbook: `resutls\jeminise.com\20260906_234129\revisions\qa_batch_032_r5\SEO_Product_Optimization_qa_batch_032_r5.xlsx`
- SHA-256: `B72388E4AADF919FB3EEAC829D7D34E7115B1300221DC9D539F7A48C31C21C92`
- Admin export: `products_export_1.csv`
- Admin export SHA-256: `729CBD9B3E0248F5D9196D999A49222DBDACA85C2F3CF8524BEEC549B57D738C`
- QA run: `20260909_134214`
- Checked at: `2026-09-09T13:39:28+07:00`
- Scope: 10 products, 70 image positions
- Batch status: `QA_PASS`
- Batch score: `96.0`

## Score Logic
- Chấm theo `seo-prompt/jeminise/prompt_qa.md` v1.0; có mở live storefront/`.js`, đối chiếu admin export, workbook r5 và contact sheet/ảnh local.
- `311-314`: lỗi `comforter` của QA r4 đã được sửa; proposed fields và 32 alt/observation liên quan đều dùng `quilt/quilt set`.
- `K3` vẫn `PARTIAL` cho toàn bộ lô vì bằng chứng nhu cầu mới ở mức `SERP_ONLY`/comparable public SERP.
- `E1` `FULL` cho 4 dòng đã sửa `r5`; `E1 PARTIAL` cho 6 dòng còn lại vì row-level `revision` vẫn là `r2` trong workbook r5. Đây là traceability `MINOR`, không phải lỗi nội dung.
- Không phát hiện `CRITICAL` hoặc `MAJOR`.

| Product | Score | Status | Images | C/M/m | Issue refs |
|---|---:|---|---:|---:|---|
| `personalized-wildlife-deer-comforter-with-buck-6532972aba-6532972aba` | 97.5 | `QA_PASS` | 8/8 | 0/0/0 | ISSUE_001 |
| `personalized-wildlife-deer-couple-in-forest-clearing-comforter-with-buck-4a23dec50b-4a23dec50b` | 97.5 | `QA_PASS` | 8/8 | 0/0/0 | ISSUE_002 |
| `personalized-wildlife-deer-couple-in-forest-comforter-0b240654a2-0b240654a2` | 97.5 | `QA_PASS` | 8/8 | 0/0/0 | ISSUE_003 |
| `personalized-wildlife-deer-couple-touching-noses-comforter-with-buck-af760d7fee-af760d7fee` | 97.5 | `QA_PASS` | 8/8 | 0/0/0 | ISSUE_004 |
| `personalized-wolf-head-framed-by-dreamcatcher-comforter-9f18113446-9f18113446` | 95.0 | `QA_PASS` | 5/5 | 0/0/1 | ISSUE_005; ISSUE_006 |
| `personalized-wolf-head-with-geometric-headdress-comforter-15e566fdc2-15e566fdc2` | 95.0 | `QA_PASS` | 5/5 | 0/0/1 | ISSUE_007; ISSUE_008 |
| `personalized-yggdrasil-tree-of-life-quilt-sets-machine-washable` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_009; ISSUE_010 |
| `personalized-yggdrasil-tree-of-life-with-ravens-quilt-sets-yggdrasil` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_011; ISSUE_012 |
| `personalized-yggdrasil-vibrant-gold-quilt-sets-machine-washable` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_013; ISSUE_014 |
| `rooster-patchwork-pattern-quilt-set-animal-bedding-chicken` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_015; ISSUE_016 |

## Findings
- `meta_description_seo`: không thấy câu bị cắt giữa từ hoặc cắt ý.
- `description_proposed_html`: không thấy câu nội bộ/template như `copy stays specific to the visible artwork`.
- `Image_Audit`: 70/70 vị trí ảnh khớp workbook, admin export và live JSON; alt mô tả đúng ảnh, không nhồi keyword.
- 10/10 sản phẩm đạt `QA_PASS` theo rubric: final_score >=85, kiểm tra đủ, không có `CRITICAL/MAJOR`.

## Limitations
- Chưa có Search Console, internal search, paid keyword volume hoặc verified review corpus; `K3` không được nâng lên `DEMAND_SUPPORTED`.
- QA không sửa workbook nguồn, không đổi `APPROVED`, không xác nhận import an toàn.
