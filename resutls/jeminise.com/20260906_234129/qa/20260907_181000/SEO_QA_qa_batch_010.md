# SEO QA — qa_batch_010

## Kết luận

- Phạm vi: **10 sản phẩm, 73/73 ảnh (100%)**; chỉ inventory position 91–100.
- Điểm lô: **67.0/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 10 QA_FAIL, 0 QA_REVISE, 0 QA_PASS.
- Phát hiện: 1 CRITICAL, 19 MAJOR, 63 MINOR, 10 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_010.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `0DA5C696AC39A2B7C6A923250474C3BBE8C2FF57DC1BF2718FC8C07C9E1B00AF`
- Ghi chú: đây là bản QA đầy đủ đã dựng lại để thay thế output rút gọn trước đó.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 91 | Custom Native American Inspired Wolf Head Inside Circular Frame Quilt | 67.5 | QA_FAIL | 0/2/4/1 |
| 92 | Custom Neon Green Soccer Player Artwork Comforter Set, Boys Room Decor | 62.5 | QA_FAIL | 1/2/7/1 |
| 93 | Custom Patriotic Semi Truck on American Flag Background Comforter | 67.5 | QA_FAIL | 0/2/6/1 |
| 94 | Custom Patriotic Semi Truck with Waving American Flag Comforter | 67.5 | QA_FAIL | 0/2/6/1 |
| 95 | Custom Patriotic Trucking Semi Truck Below Waving American Flag Comforter | 67.5 | QA_FAIL | 0/2/6/1 |
| 96 | Custom Photo Quilt with Bicycle and Flowers | 67.5 | QA_FAIL | 0/1/6/1 |
| 97 | Custom Purple Butterflies Text Blanket, Machine Washable - D7b | 67.5 | QA_FAIL | 0/2/7/1 |
| 98 | Custom Purple Butterfly Christian Blanket, Machine Washable - D15 | 67.5 | QA_FAIL | 0/2/7/1 |
| 99 | Custom Purple Floral Christian Cross Blanket, Machine Washable - D2 | 67.5 | QA_FAIL | 0/2/7/1 |
| 100 | Custom Purple Floral Cross with Butterflies Blanket, Machi - Design 14 | 67.5 | QA_FAIL | 0/2/7/1 |

## Lỗi ưu tiên

1. **CRITICAL — product 92:** Claim is materially unsafe until the live purchase flow exposes matching input controls. Đề xuất: Remove/soften unsupported personalization claim, or provide admin/live proof that matching controls work.
2. **MAJOR — product 91:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
3. **MAJOR — product 91:** Batch cluster contains several very similar products, so keyword/intent differentiation must be sharper. Đề xuất: Add unique visible motif modifiers and avoid assigning the same generic product keyword to multiple URLs.
4. **MAJOR — product 92:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
5. **MAJOR — product 92:** Batch cluster contains several very similar products, so keyword/intent differentiation must be sharper. Đề xuất: Add unique visible motif modifiers and avoid assigning the same generic product keyword to multiple URLs.
6. **MAJOR — product 93:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
7. **MAJOR — product 93:** Batch cluster contains several very similar products, so keyword/intent differentiation must be sharper. Đề xuất: Add unique visible motif modifiers and avoid assigning the same generic product keyword to multiple URLs.
8. **MAJOR — product 94:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.

## SERP và keyword

Mỗi sản phẩm đã được ghi lại tối thiểu primary keyword và một comparator gần nhất trong `serp_evidence.json`. Evidence chỉ là US public SERP/product comparables, không có claim volume.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; không suy ra admin SEO fields hoặc alt hiện tại ngoài storefront/product.js.
- Đã dùng workbook snapshot, live_source_comparison, product.js/static HTML, customizer_audit và ảnh tải trực tiếp trong thư mục QA run.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại kiểm tra cấu trúc.
- Chưa QA batch tiếp theo sau qa_batch_010. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_181000\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_181000\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_181000\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_181000`
