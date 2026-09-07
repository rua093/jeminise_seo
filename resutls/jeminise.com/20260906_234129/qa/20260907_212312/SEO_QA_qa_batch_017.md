# SEO QA — qa_batch_017

## Kết luận

- Phạm vi: **10 sản phẩm, 62/62 ảnh (100%)**; chỉ inventory position 161–170.
- Điểm lô: **62.5/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 10 QA_FAIL, 0 QA_REVISE, 0 QA_PASS.
- Phát hiện: 0 CRITICAL, 21 MAJOR, 61 MINOR, 10 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_017.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `65F8C48D63E55BA5F6B311B33EF5ED1FBFEAB99A148B84E45592C180C8CD9132`
- Ghi chú: đây là bản QA độc lập đầy đủ theo `prompt_qa.md`, revision r1.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 161 | Norse Mythology Raven Within Celtic Knot Frame Quilt | 62.5 | QA_FAIL | 0/2/7/1 |
| 162 | Norse Mythology Viking Shield with Crossed Axes Quilt | 62.5 | QA_FAIL | 0/3/7/1 |
| 163 | Owl Patchwork Pattern Quilt Set, Animals Quilt Set Style, Owl Nights | 62.5 | QA_FAIL | 0/2/7/1 |
| 164 | Pamnest Twin Halloween Comforter Set with Sheets, 5 Pieces - Design 01 | 62.5 | QA_FAIL | 0/2/6/1 |
| 165 | Pamnest Twin Halloween Comforter Set with Sheets, 5 Pieces - Design 02 | 62.5 | QA_FAIL | 0/2/4/1 |
| 166 | Pamnest Twin Halloween Comforter Set with Sheets, 5 Pieces - Design 03 | 62.5 | QA_FAIL | 0/2/6/1 |
| 167 | Pamnest Twin Halloween Comforter Set with Sheets, 5 Pieces - Design 04 | 62.5 | QA_FAIL | 0/2/6/1 |
| 168 | Pamnest Twin Halloween Comforter Set with Sheets, 5 Pieces - Design 05 | 62.5 | QA_FAIL | 0/2/6/1 |
| 169 | Pamnest Twin Halloween Comforter Set with Sheets, 5 Pieces - Design 06 | 62.5 | QA_FAIL | 0/2/6/1 |
| 170 | Pamnest Twin Halloween Comforter Set with Sheets, 5 Pieces - Design 07 | 62.5 | QA_FAIL | 0/2/6/1 |

## Lỗi ưu tiên

1. **MAJOR — product 161:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
2. **MAJOR — product 161:** A publishable SEO title/H1 pair is incomplete when the H1 field is blank. Đề xuất: Write one English H1 that matches the verified product type and visible design; do not copy unsupported claims.
3. **MAJOR — product 162:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
4. **MAJOR — product 162:** A publishable SEO title/H1 pair is incomplete when the H1 field is blank. Đề xuất: Write one English H1 that matches the verified product type and visible design; do not copy unsupported claims.
5. **MAJOR — product 162:** Batch cluster contains several very similar products, so keyword/intent differentiation must be sharper. Đề xuất: Add unique visible motif modifiers and avoid assigning the same generic product keyword to multiple URLs.
6. **MAJOR — product 163:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
7. **MAJOR — product 163:** A publishable SEO title/H1 pair is incomplete when the H1 field is blank. Đề xuất: Write one English H1 that matches the verified product type and visible design; do not copy unsupported claims.
8. **MAJOR — product 164:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.

## SERP và keyword

Mỗi sản phẩm đã được ghi lại tối thiểu primary keyword và một comparator gần nhất trong `serp_evidence.json`. Evidence chỉ là US public SERP/product comparables, không có claim volume.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; không suy ra admin SEO fields hoặc alt hiện tại ngoài storefront/product.js.
- Đã dùng workbook snapshot, live_source_comparison, product.js/static HTML, customizer_audit và ảnh tải trực tiếp trong thư mục QA run.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại kiểm tra cấu trúc, công thức và bố cục.
- Chưa QA batch tiếp theo sau qa_batch_017. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_212312\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_212312\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_212312\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_212312`
