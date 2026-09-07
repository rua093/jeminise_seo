# SEO QA — qa_batch_013

## Kết luận

- Phạm vi: **10 sản phẩm, 62/62 ảnh (100%)**; chỉ inventory position 121–130.
- Điểm lô: **67.0/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 10 QA_FAIL, 0 QA_REVISE, 0 QA_PASS.
- Phát hiện: 1 CRITICAL, 19 MAJOR, 52 MINOR, 10 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_013.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `365DE89A3E4E2CC44D50F426944B4728DDE44EE3ACB68398663F6EED614D2700`
- Ghi chú: đây là bản QA đầy đủ đã dựng lại để thay thế output rút gọn trước đó.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 121 | Custom Trucking Semi Truck Breaking Through Stone Wall Vintage Comforter | 67.5 | QA_FAIL | 0/2/6/1 |
| 122 | Custom Trucking Semi Truck Centered on Chevron Background Comforter | 67.5 | QA_FAIL | 0/2/6/1 |
| 123 | Custom Trucking Semi Truck on Metallic Mesh Background Comforter | 67.5 | QA_FAIL | 0/2/6/1 |
| 124 | Custom Trucking Semi Truck on Textured Metallic Background Comforter | 67.5 | QA_FAIL | 0/2/6/1 |
| 125 | Custom Trucking Semi Truck Under Starry Night Sky Comforter | 67.5 | QA_FAIL | 0/2/6/1 |
| 126 | Custom Two Cardinals on Holly Branches Quilt | 67.5 | QA_FAIL | 0/1/6/1 |
| 127 | Custom Wildlife Wolf Head Close Up Portrait Quilt | 67.5 | QA_FAIL | 0/2/4/1 |
| 128 | Custom Wildlife Wolf Head Close Up with Geometric Quilt | 62.5 | QA_FAIL | 1/2/4/1 |
| 129 | Custom Wildlife Wolf Head with Feathers in Profile Quilt | 67.5 | QA_FAIL | 0/2/4/1 |
| 130 | Custom Wolf Head Centered in Dreamcatcher Quilt | 67.5 | QA_FAIL | 0/2/4/1 |

## Lỗi ưu tiên

1. **CRITICAL — product 128:** Claim is materially unsafe until the live purchase flow exposes matching input controls. Đề xuất: Remove/soften unsupported personalization claim, or provide admin/live proof that matching controls work.
2. **MAJOR — product 121:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
3. **MAJOR — product 121:** Batch cluster contains several very similar products, so keyword/intent differentiation must be sharper. Đề xuất: Add unique visible motif modifiers and avoid assigning the same generic product keyword to multiple URLs.
4. **MAJOR — product 122:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
5. **MAJOR — product 122:** Batch cluster contains several very similar products, so keyword/intent differentiation must be sharper. Đề xuất: Add unique visible motif modifiers and avoid assigning the same generic product keyword to multiple URLs.
6. **MAJOR — product 123:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
7. **MAJOR — product 123:** Batch cluster contains several very similar products, so keyword/intent differentiation must be sharper. Đề xuất: Add unique visible motif modifiers and avoid assigning the same generic product keyword to multiple URLs.
8. **MAJOR — product 124:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.

## SERP và keyword

Mỗi sản phẩm đã được ghi lại tối thiểu primary keyword và một comparator gần nhất trong `serp_evidence.json`. Evidence chỉ là US public SERP/product comparables, không có claim volume.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; không suy ra admin SEO fields hoặc alt hiện tại ngoài storefront/product.js.
- Đã dùng workbook snapshot, live_source_comparison, product.js/static HTML, customizer_audit và ảnh tải trực tiếp trong thư mục QA run.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại kiểm tra cấu trúc.
- Chưa QA batch tiếp theo sau qa_batch_013. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_182000\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_182000\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_182000\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_182000`
