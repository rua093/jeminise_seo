# SEO QA — qa_batch_012

## Kết luận

- Phạm vi: **10 sản phẩm, 72/72 ảnh (100%)**; chỉ inventory position 111–120.
- Điểm lô: **66.2/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 10 QA_FAIL, 0 QA_REVISE, 0 QA_PASS.
- Phát hiện: 3 CRITICAL, 17 MAJOR, 61 MINOR, 10 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_012.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `6D76C36179A27F454621DBC098D3621663526386146A212D6C326832CC7E9D05`
- Ghi chú: đây là bản QA đầy đủ đã dựng lại để thay thế output rút gọn trước đó.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 111 | Custom Soccer Player Custom Name Number Comforter Set, Boys Room Decor | 62.5 | QA_FAIL | 1/2/7/1 |
| 112 | Custom Soccer Player Goal Comforter Set, Machine Washable - Design 8 | 65.0 | QA_FAIL | 1/2/6/1 |
| 113 | Custom Soccer Player Paint Splatter Comforter Set, Machine Washable | 62.5 | QA_FAIL | 1/2/7/1 |
| 114 | Custom Softball Comforter Set for Girls – Personalized Name & Number | 67.5 | QA_FAIL | 0/1/6/1 |
| 115 | Custom Sports and Patriotism Football on American Flag Background Comforter | 67.5 | QA_FAIL | 0/2/5/1 |
| 116 | Custom Transportation Semi Truck on Geometric Background Comforter | 67.5 | QA_FAIL | 0/2/6/1 |
| 117 | Custom Tree of Life Quilt Set Celtic Yggdrasil Bedding | 67.5 | QA_FAIL | 0/1/6/1 |
| 118 | Custom Tree of Life Quilt Set Celtic Yggdrasil Bedding for Bedroom Decor | 67.5 | QA_FAIL | 0/1/6/1 |
| 119 | Custom Trucking Semi Truck Breaking Through Stone Wall Comforter design 1 | 67.5 | QA_FAIL | 0/2/6/1 |
| 120 | Custom Trucking Semi Truck Breaking Through Stone Wall Comforter design 2 | 67.5 | QA_FAIL | 0/2/6/1 |

## Lỗi ưu tiên

1. **CRITICAL — product 111:** Claim is materially unsafe until the live purchase flow exposes matching input controls. Đề xuất: Remove/soften unsupported personalization claim, or provide admin/live proof that matching controls work.
2. **CRITICAL — product 112:** Claim is materially unsafe until the live purchase flow exposes matching input controls. Đề xuất: Remove/soften unsupported personalization claim, or provide admin/live proof that matching controls work.
3. **CRITICAL — product 113:** Claim is materially unsafe until the live purchase flow exposes matching input controls. Đề xuất: Remove/soften unsupported personalization claim, or provide admin/live proof that matching controls work.
4. **MAJOR — product 111:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
5. **MAJOR — product 111:** Batch cluster contains several very similar products, so keyword/intent differentiation must be sharper. Đề xuất: Add unique visible motif modifiers and avoid assigning the same generic product keyword to multiple URLs.
6. **MAJOR — product 112:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
7. **MAJOR — product 112:** Batch cluster contains several very similar products, so keyword/intent differentiation must be sharper. Đề xuất: Add unique visible motif modifiers and avoid assigning the same generic product keyword to multiple URLs.
8. **MAJOR — product 113:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.

## SERP và keyword

Mỗi sản phẩm đã được ghi lại tối thiểu primary keyword và một comparator gần nhất trong `serp_evidence.json`. Evidence chỉ là US public SERP/product comparables, không có claim volume.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; không suy ra admin SEO fields hoặc alt hiện tại ngoài storefront/product.js.
- Đã dùng workbook snapshot, live_source_comparison, product.js/static HTML, customizer_audit và ảnh tải trực tiếp trong thư mục QA run.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại kiểm tra cấu trúc.
- Chưa QA batch tiếp theo sau qa_batch_012. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_181200\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_181200\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_181200\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_181200`
