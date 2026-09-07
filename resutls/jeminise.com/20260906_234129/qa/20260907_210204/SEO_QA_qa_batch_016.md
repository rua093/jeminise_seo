# SEO QA — qa_batch_016

## Kết luận

- Phạm vi: **10 sản phẩm, 71/71 ảnh (100%)**; chỉ inventory position 151–160.
- Điểm lô: **62.5/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 10 QA_FAIL, 0 QA_REVISE, 0 QA_PASS.
- Phát hiện: 0 CRITICAL, 31 MAJOR, 61 MINOR, 10 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_016.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `16735BD1FE1D51206D00E8B326215ACC9305ADF149AFCF477CA059CA7E8A9087`
- Ghi chú: đây là bản QA độc lập đầy đủ theo `prompt_qa.md`, revision r1.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 151 | Giraffe Patchwork Style Printed Quilt, Animal Patchwork Bedding | 62.5 | QA_FAIL | 0/3/6/1 |
| 152 | Horse Patchwork Pattern Quilt Set, Vintage Animal Quilt Printed | 62.5 | QA_FAIL | 0/3/6/1 |
| 153 | Horse Patchwork Pattern Quilt, Animal Farmhouse Quilt Set Patchwork | 62.5 | QA_FAIL | 0/3/6/1 |
| 154 | I Am Who He Says I Am Christian Scripture Blanket – Inspirational | 62.5 | QA_FAIL | 0/4/7/1 |
| 155 | Nature and Mythology Tree of Life Quilt Set With Twisted Tree Trunk | 62.5 | QA_FAIL | 0/3/6/1 |
| 156 | Nature Tree of Life Quilt Set with Blooming Tree and Bedroom And Exposed Roots | 62.5 | QA_FAIL | 0/3/6/1 |
| 157 | Nature Tree of Life Quilt Set with Stained Glass Pattern And Bedroom | 62.5 | QA_FAIL | 0/3/6/1 |
| 158 | Norse Mythology Mjolnir with Skull Within Circular Celtic Quilt | 62.5 | QA_FAIL | 0/3/6/1 |
| 159 | Norse Mythology Quilt with Ravens | 62.5 | QA_FAIL | 0/3/6/1 |
| 160 | Norse Mythology Raven Centered Within Runic Circle Comforter | 62.5 | QA_FAIL | 0/3/6/1 |

## Lỗi ưu tiên

1. **MAJOR — product 151:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
2. **MAJOR — product 151:** A publishable SEO title/H1 pair is incomplete when the H1 field is blank. Đề xuất: Write one English H1 that matches the verified product type and visible design; do not copy unsupported claims.
3. **MAJOR — product 151:** Generic audience wording can be irrelevant or contradictory to the visible motif and is not evidence-led storefront copy. Đề xuất: Replace it with concise English customer-facing copy tied only to the verified motif, product type, size/pillowcase choices and supported customizer control.
4. **MAJOR — product 152:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
5. **MAJOR — product 152:** A publishable SEO title/H1 pair is incomplete when the H1 field is blank. Đề xuất: Write one English H1 that matches the verified product type and visible design; do not copy unsupported claims.
6. **MAJOR — product 152:** Generic audience wording can be irrelevant or contradictory to the visible motif and is not evidence-led storefront copy. Đề xuất: Replace it with concise English customer-facing copy tied only to the verified motif, product type, size/pillowcase choices and supported customizer control.
7. **MAJOR — product 153:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
8. **MAJOR — product 153:** A publishable SEO title/H1 pair is incomplete when the H1 field is blank. Đề xuất: Write one English H1 that matches the verified product type and visible design; do not copy unsupported claims.

## SERP và keyword

Mỗi sản phẩm đã được ghi lại tối thiểu primary keyword và một comparator gần nhất trong `serp_evidence.json`. Evidence chỉ là US public SERP/product comparables, không có claim volume.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; không suy ra admin SEO fields hoặc alt hiện tại ngoài storefront/product.js.
- Đã dùng workbook snapshot, live_source_comparison, product.js/static HTML, customizer_audit và ảnh tải trực tiếp trong thư mục QA run.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại kiểm tra cấu trúc, công thức và bố cục.
- Chưa QA batch tiếp theo sau qa_batch_016. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_210204\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_210204\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_210204\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_210204`
