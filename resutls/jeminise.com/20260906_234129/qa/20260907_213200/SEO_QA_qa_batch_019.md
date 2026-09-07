# SEO QA — qa_batch_019

## Kết luận

- Phạm vi: **10 sản phẩm, 40/40 ảnh (100%)**; chỉ inventory position 181–190.
- Điểm lô: **62.5/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 10 QA_FAIL, 0 QA_REVISE, 0 QA_PASS.
- Phát hiện: 0 CRITICAL, 20 MAJOR, 40 MINOR, 10 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_019.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `DFC6E245CAD3637D002D2DE0D2A60EFD9D81574462722DAB0043D63444EF8CA1`
- Note: this is the complete independent QA following `prompt_qa.md`, revision r1.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 181 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 06 | 62.5 | QA_FAIL | 0/2/4/1 |
| 182 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 07 | 62.5 | QA_FAIL | 0/2/4/1 |
| 183 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 08 | 62.5 | QA_FAIL | 0/2/4/1 |
| 184 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 09 | 62.5 | QA_FAIL | 0/2/4/1 |
| 185 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 10 | 62.5 | QA_FAIL | 0/2/4/1 |
| 186 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 11 | 62.5 | QA_FAIL | 0/2/4/1 |
| 187 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 12 | 62.5 | QA_FAIL | 0/2/4/1 |
| 188 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 13 | 62.5 | QA_FAIL | 0/2/4/1 |
| 189 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 14 | 62.5 | QA_FAIL | 0/2/4/1 |
| 190 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 15 | 62.5 | QA_FAIL | 0/2/4/1 |

## Lỗi ưu tiên

1. **MAJOR — product 181:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
2. **MAJOR — product 181:** A publishable SEO title/H1 pair is incomplete when the H1 field is blank. Đề xuất: Write one English H1 that matches the verified product type and visible design; do not copy unsupported claims.
3. **MAJOR — product 182:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
4. **MAJOR — product 182:** A publishable SEO title/H1 pair is incomplete when the H1 field is blank. Đề xuất: Write one English H1 that matches the verified product type and visible design; do not copy unsupported claims.
5. **MAJOR — product 183:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
6. **MAJOR — product 183:** A publishable SEO title/H1 pair is incomplete when the H1 field is blank. Đề xuất: Write one English H1 that matches the verified product type and visible design; do not copy unsupported claims.
7. **MAJOR — product 184:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
8. **MAJOR — product 184:** A publishable SEO title/H1 pair is incomplete when the H1 field is blank. Đề xuất: Write one English H1 that matches the verified product type and visible design; do not copy unsupported claims.

## SERP và keyword

Mỗi sản phẩm đã được ghi lại tối thiểu primary keyword và một comparator gần nhất trong `serp_evidence.json`. Evidence chỉ là US public SERP/product comparables, không có claim volume.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; không suy ra admin SEO fields hoặc alt hiện tại ngoài storefront/product.js.
- Đã dùng workbook snapshot, live_source_comparison, product.js/static HTML, customizer_audit và ảnh tải trực tiếp trong thư mục QA run.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA has exactly five sheets, auditable formulas, filters, frozen headers, wrapped text and hyperlinks; structure, formulas and layout were reopened and checked.
- Chưa QA batch tiếp theo sau qa_batch_019. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_213200\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_213200\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_213200\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_213200`
