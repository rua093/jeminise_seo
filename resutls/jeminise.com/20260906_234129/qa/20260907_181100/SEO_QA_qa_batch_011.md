# SEO QA — qa_batch_011

## Kết luận

- Phạm vi: **10 sản phẩm, 77/77 ảnh (100%)**; chỉ inventory position 101–110.
- Điểm lô: **63.5/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 10 QA_FAIL, 0 QA_REVISE, 0 QA_PASS.
- Phát hiện: 9 CRITICAL, 20 MAJOR, 65 MINOR, 10 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_011.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `2F0832604D71BC81D1EEC3DCC1F1CDAD85B8340EBCD3808B6D74BDAAE05ECF49`
- Ghi chú: đây là bản QA đầy đủ đã dựng lại để thay thế output rút gọn trước đó.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 101 | Custom Semi Truck with American Flag Background Comforter | 67.5 | QA_FAIL | 0/2/6/1 |
| 102 | Custom Soccer Ball Cleats Personalized Comforter Set, Machine Washable | 62.5 | QA_FAIL | 1/2/7/1 |
| 103 | Custom Soccer Ball Goal Net Comforter Set, Machine Washabl - Design 13 | 65.0 | QA_FAIL | 1/2/6/1 |
| 104 | Custom Soccer Ball Goal Personalized Bedding Comforter Set - A4 | 62.5 | QA_FAIL | 1/2/6/1 |
| 105 | Custom Soccer Ball Net Comforter Set, Machine Washable - A2 | 62.5 | QA_FAIL | 1/2/7/1 |
| 106 | Custom Soccer Ball Water Splashes Comforter Set, Machine Washable | 65.0 | QA_FAIL | 1/2/6/1 |
| 107 | Custom Soccer Ball with Lightning Splash Comforter Set - A11 | 62.5 | QA_FAIL | 1/2/7/1 |
| 108 | Custom Soccer Ball with Number 07 Comforter Set, Machine W - Design 17 | 62.5 | QA_FAIL | 1/2/7/1 |
| 109 | Custom Soccer Goal Splash Bedding Comforter Set, Machine Washable | 62.5 | QA_FAIL | 1/2/6/1 |
| 110 | Custom Soccer Net with Number 07 Comforter Set, Machine Washable - A7 | 62.5 | QA_FAIL | 1/2/7/1 |

## Lỗi ưu tiên

1. **CRITICAL — product 102:** Claim is materially unsafe until the live purchase flow exposes matching input controls. Đề xuất: Remove/soften unsupported personalization claim, or provide admin/live proof that matching controls work.
2. **CRITICAL — product 103:** Claim is materially unsafe until the live purchase flow exposes matching input controls. Đề xuất: Remove/soften unsupported personalization claim, or provide admin/live proof that matching controls work.
3. **CRITICAL — product 104:** Claim is materially unsafe until the live purchase flow exposes matching input controls. Đề xuất: Remove/soften unsupported personalization claim, or provide admin/live proof that matching controls work.
4. **CRITICAL — product 105:** Claim is materially unsafe until the live purchase flow exposes matching input controls. Đề xuất: Remove/soften unsupported personalization claim, or provide admin/live proof that matching controls work.
5. **CRITICAL — product 106:** Claim is materially unsafe until the live purchase flow exposes matching input controls. Đề xuất: Remove/soften unsupported personalization claim, or provide admin/live proof that matching controls work.
6. **CRITICAL — product 107:** Claim is materially unsafe until the live purchase flow exposes matching input controls. Đề xuất: Remove/soften unsupported personalization claim, or provide admin/live proof that matching controls work.
7. **CRITICAL — product 108:** Claim is materially unsafe until the live purchase flow exposes matching input controls. Đề xuất: Remove/soften unsupported personalization claim, or provide admin/live proof that matching controls work.
8. **CRITICAL — product 109:** Claim is materially unsafe until the live purchase flow exposes matching input controls. Đề xuất: Remove/soften unsupported personalization claim, or provide admin/live proof that matching controls work.

## SERP và keyword

Mỗi sản phẩm đã được ghi lại tối thiểu primary keyword và một comparator gần nhất trong `serp_evidence.json`. Evidence chỉ là US public SERP/product comparables, không có claim volume.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; không suy ra admin SEO fields hoặc alt hiện tại ngoài storefront/product.js.
- Đã dùng workbook snapshot, live_source_comparison, product.js/static HTML, customizer_audit và ảnh tải trực tiếp trong thư mục QA run.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại kiểm tra cấu trúc.
- Chưa QA batch tiếp theo sau qa_batch_011. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_181100\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_181100\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_181100\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_181100`
