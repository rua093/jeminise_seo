# SEO QA — qa_batch_020

## Kết luận

- Phạm vi: **10 sản phẩm, 46/46 ảnh (100%)**; chỉ inventory position 191–200.
- Điểm lô: **62.5/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 10 QA_FAIL, 0 QA_REVISE, 0 QA_PASS.
- Phát hiện: 0 CRITICAL, 20 MAJOR, 44 MINOR, 10 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_020.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `513379577FAB118478068CC65AEFDA27D16A01A945472683ED57FEAA1C80D5C4`
- Note: this is the complete independent QA following `prompt_qa.md`, revision r1.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 191 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 16 | 62.5 | QA_FAIL | 0/2/6/1 |
| 192 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 17 | 62.5 | QA_FAIL | 0/2/6/1 |
| 193 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 18 | 62.5 | QA_FAIL | 0/2/4/1 |
| 194 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 19 | 62.5 | QA_FAIL | 0/2/4/1 |
| 195 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 20 | 62.5 | QA_FAIL | 0/2/4/1 |
| 196 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 21 | 62.5 | QA_FAIL | 0/2/4/1 |
| 197 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 22 | 62.5 | QA_FAIL | 0/2/4/1 |
| 198 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 23 | 62.5 | QA_FAIL | 0/2/4/1 |
| 199 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 24 | 62.5 | QA_FAIL | 0/2/4/1 |
| 200 | Personalized Baseball Bedding Full Size Flag Custom Name D - Design 25 | 62.5 | QA_FAIL | 0/2/4/1 |

## Lỗi ưu tiên

1. **MAJOR — product 191:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
2. **MAJOR — product 191:** A publishable SEO title/H1 pair is incomplete when the H1 field is blank. Đề xuất: Write one English H1 that matches the verified product type and visible design; do not copy unsupported claims.
3. **MAJOR — product 192:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
4. **MAJOR — product 192:** A publishable SEO title/H1 pair is incomplete when the H1 field is blank. Đề xuất: Write one English H1 that matches the verified product type and visible design; do not copy unsupported claims.
5. **MAJOR — product 193:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
6. **MAJOR — product 193:** A publishable SEO title/H1 pair is incomplete when the H1 field is blank. Đề xuất: Write one English H1 that matches the verified product type and visible design; do not copy unsupported claims.
7. **MAJOR — product 194:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
8. **MAJOR — product 194:** A publishable SEO title/H1 pair is incomplete when the H1 field is blank. Đề xuất: Write one English H1 that matches the verified product type and visible design; do not copy unsupported claims.

## SERP và keyword

Mỗi sản phẩm đã được ghi lại tối thiểu primary keyword và một comparator gần nhất trong `serp_evidence.json`. Evidence chỉ là US public SERP/product comparables, không có claim volume.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; không suy ra admin SEO fields hoặc alt hiện tại ngoài storefront/product.js.
- Đã dùng workbook snapshot, live_source_comparison, product.js/static HTML, customizer_audit và ảnh tải trực tiếp trong thư mục QA run.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA has exactly five sheets, auditable formulas, filters, frozen headers, wrapped text and hyperlinks; structure, formulas and layout were reopened and checked.
- Chưa QA batch tiếp theo sau qa_batch_020. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_214100\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_214100\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_214100\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_214100`
