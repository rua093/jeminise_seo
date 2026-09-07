# SEO QA — qa_batch_014

## Kết luận

- Phạm vi: **10 sản phẩm, 68/68 ảnh (100%)**; chỉ inventory position 131–140.
- Điểm lô: **67.5/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 10 QA_FAIL, 0 QA_REVISE, 0 QA_PASS.
- Phát hiện: 0 CRITICAL, 20 MAJOR, 58 MINOR, 10 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_014.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `E4C5566A8039580B94A70655E65ABF4C34932F9624B9038EE5F50C5840898484`
- Ghi chú: đây là bản QA độc lập đầy đủ cho lô qa_batch_014 theo đúng rubric prompt_qa.md.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 131 | Custom Wolf Wearing Headdress Centered Close Up Quilt with Wolf Head | 67.5 | QA_FAIL | 0/2/4/1 |
| 132 | Decorative Softball Bedding Set – Bright Teal Yellow Sports Comforter | 67.5 | QA_FAIL | 0/2/6/1 |
| 133 | Desert Cactus Printed Quilt Set, Vibrant Cactus & Flower Quilt | 67.5 | QA_FAIL | 0/2/6/1 |
| 134 | Dragonfly Couple Quilt Set, Dragonflies Patchwork Style Quilt Set | 67.5 | QA_FAIL | 0/2/6/1 |
| 135 | Dragonfly Patchwork Pattern Quilt Set, Animal Style, Dragonfly Quilt | 67.5 | QA_FAIL | 0/2/6/1 |
| 136 | Dragonfly Patchwork Quilt Set, Animal Patchwork Bedset - Dragonfly | 67.5 | QA_FAIL | 0/2/6/1 |
| 137 | Dragonfly Pattern Quilt Set, Dragonflies Vintage Quilt Set Patchwork | 67.5 | QA_FAIL | 0/2/6/1 |
| 138 | Elephant Patchwork Pattern Quilt Set, Animal Farmhouse Quilt Set | 67.5 | QA_FAIL | 0/2/6/1 |
| 139 | Elephant Quilt Set, Animal Patchwork Printed Quilt Bedset, Patchwork | 67.5 | QA_FAIL | 0/2/6/1 |
| 140 | Fantasy Art Fantasy Dragon Quilt Set with Dragon Scales And Bedroom Setting | 67.5 | QA_FAIL | 0/2/6/1 |

## Lỗi ưu tiên

1. **MAJOR — product 131:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
2. **MAJOR — product 131:** Batch cluster contains several very similar products, so keyword/intent differentiation must be sharper. Đề xuất: Add unique visible motif modifiers and avoid assigning the same generic product keyword to multiple URLs.
3. **MAJOR — product 132:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
4. **MAJOR — product 132:** Batch cluster contains several very similar products, so keyword/intent differentiation must be sharper. Đề xuất: Add unique visible motif modifiers and avoid assigning the same generic product keyword to multiple URLs.
5. **MAJOR — product 133:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
6. **MAJOR — product 133:** Batch cluster contains several very similar products, so keyword/intent differentiation must be sharper. Đề xuất: Add unique visible motif modifiers and avoid assigning the same generic product keyword to multiple URLs.
7. **MAJOR — product 134:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
8. **MAJOR — product 134:** Batch cluster contains several very similar products, so keyword/intent differentiation must be sharper. Đề xuất: Add unique visible motif modifiers and avoid assigning the same generic product keyword to multiple URLs.

## SERP và keyword

Mỗi sản phẩm đã được ghi lại tối thiểu primary keyword và một comparator gần nhất trong `serp_evidence.json`. Evidence chỉ là US public SERP/product comparables, không có claim volume.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; không suy ra admin SEO fields hoặc alt hiện tại ngoài storefront/product.js.
- Đã dùng workbook snapshot, live_source_comparison, product.js/static HTML, customizer_audit và ảnh tải trực tiếp trong thư mục QA run.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; render kiểm tra nằm trong `rendered_sheets`.
- Chưa QA batch tiếp theo sau qa_batch_014. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_204400\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_204400\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_204400\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_204400`
