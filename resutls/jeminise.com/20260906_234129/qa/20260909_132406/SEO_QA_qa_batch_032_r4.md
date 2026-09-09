# SEO QA qa_batch_032_r4

- Source workbook: `resutls\jeminise.com\20260906_234129\revisions\qa_batch_032_r4\SEO_Product_Optimization_qa_batch_032_r4.xlsx`
- SHA-256: `4CF64DF7E3E10ABB5446DC8CB91A0A32241FB3E9FE8D01A1DF6314659922A34B`
- Admin export: `products_export_1.csv`
- Admin export SHA-256: `729CBD9B3E0248F5D9196D999A49222DBDACA85C2F3CF8524BEEC549B57D738C`
- QA run: `20260909_132406`
- Checked at: `2026-09-09T13:28:00+07:00`
- Scope: 10 products, 70 image positions
- Batch status: `QA_REVISE`
- Batch score: `86.2`

## Score Logic
- Chấm theo `seo-prompt/jeminise/prompt_qa.md` v1.0, có mở live storefront/`.js`, đối chiếu `products_export_1.csv`, workbook r4 và contact sheet/ảnh local.
- `311-314` bị `MAJOR` vì public SEO fields và alt đang dùng `comforter/comforter set`, trong khi admin `Type`, option size và admin SEO baseline neo về `Quilt`; cần merchant xác nhận hoặc sửa đồng nhất về `Quilt/Quilt Set`.
- `K3` bị `PARTIAL` cho toàn bộ lô vì bằng chứng nhu cầu mới ở mức public SERP/comparable shopping results, chưa có Search Console/internal search/paid keyword data.
- `E1` bị `PARTIAL` cho toàn bộ lô vì source workbook là `qa_batch_032_r4` nhưng row-level `revision` vẫn ghi `r2`.

| Product | Score | Status | Images | C/M/m | Issue refs |
|---|---:|---|---:|---:|---|
| `personalized-wildlife-deer-comforter-with-buck-6532972aba-6532972aba` | 73.0 | `QA_REVISE` | 8/8 | 0/1/1 | ISSUE_001; ISSUE_002; ISSUE_003 |
| `personalized-wildlife-deer-couple-in-forest-clearing-comforter-with-buck-4a23dec50b-4a23dec50b` | 73.0 | `QA_REVISE` | 8/8 | 0/1/1 | ISSUE_004; ISSUE_005; ISSUE_006 |
| `personalized-wildlife-deer-couple-in-forest-comforter-0b240654a2-0b240654a2` | 73.0 | `QA_REVISE` | 8/8 | 0/1/1 | ISSUE_007; ISSUE_008; ISSUE_009 |
| `personalized-wildlife-deer-couple-touching-noses-comforter-with-buck-af760d7fee-af760d7fee` | 73.0 | `QA_REVISE` | 8/8 | 0/1/1 | ISSUE_010; ISSUE_011; ISSUE_012 |
| `personalized-wolf-head-framed-by-dreamcatcher-comforter-9f18113446-9f18113446` | 95.0 | `QA_PASS` | 5/5 | 0/0/1 | ISSUE_013; ISSUE_014 |
| `personalized-wolf-head-with-geometric-headdress-comforter-15e566fdc2-15e566fdc2` | 95.0 | `QA_PASS` | 5/5 | 0/0/1 | ISSUE_015; ISSUE_016 |
| `personalized-yggdrasil-tree-of-life-quilt-sets-machine-washable` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_017; ISSUE_018 |
| `personalized-yggdrasil-tree-of-life-with-ravens-quilt-sets-yggdrasil` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_019; ISSUE_020 |
| `personalized-yggdrasil-vibrant-gold-quilt-sets-machine-washable` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_021; ISSUE_022 |
| `rooster-patchwork-pattern-quilt-set-animal-bedding-chicken` | 95.0 | `QA_PASS` | 7/7 | 0/0/1 | ISSUE_023; ISSUE_024 |

## Findings
- `meta_description_seo`: không thấy câu bị cắt giữa từ hoặc cắt ý; các meta là câu hoàn chỉnh.
- `description_proposed_html`: không thấy câu nội bộ/template như `copy stays specific to the visible artwork`.
- `Image_Audit`: 70/70 vị trí ảnh khớp số lượng live JSON và admin export.
- 6 sản phẩm `315-320` đạt `QA_PASS`; 4 sản phẩm deer `311-314` cần revision trước khi gửi pass do lỗi nhất quán loại sản phẩm.

## Priority Fix
1. Sửa `311-314` để dùng `Quilt`/`Quilt Set` đồng nhất trong `title_proposed`, `meta_title_seo`, `meta_description_seo`, `description_proposed_html` và tất cả `alt_proposed`, hoặc cung cấp xác nhận merchant rằng 4 sản phẩm này phải market là `comforter`.
2. Cập nhật metadata `revision` trong row lên đúng `r4/r5` ở bản sửa tiếp theo để traceability rõ.

## Limitations
- Chưa có Search Console, internal search, paid keyword volume hoặc review corpus xác minh nhu cầu; chỉ dùng public SERP/comparable shopping pages.
- QA không sửa workbook nguồn và không đổi trạng thái `APPROVED`.
