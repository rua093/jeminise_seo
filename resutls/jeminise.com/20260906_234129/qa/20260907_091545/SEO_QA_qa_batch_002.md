# SEO QA — qa_batch_002

## Kết luận

- Phạm vi: **10 sản phẩm, 64/64 ảnh (100%)**; chỉ inventory position 11–20.
- Điểm lô: **59.8/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 7 QA_FAIL, 3 QA_REVISE, 0 QA_PASS.
- Phát hiện: 7 CRITICAL, 51 MAJOR, 20 MINOR, 1 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_034.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `609EF38D46896F234EA7AAEEC74F393267003D20D6DE8C5BA9BD21BFB19675EE`

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 11 | Christian Bedding Set – Inspirational Bible Verse Comforte - Design 02 | 46.0 | QA_FAIL | 1/6/4/0 |
| 12 | Christian Bedding Set – Inspirational Bible Verse Comforte - Design 03 | 45.1 | QA_FAIL | 1/7/2/0 |
| 13 | Christian Bedding Set – Inspirational Bible Verse Comforte - Design 04 | 45.1 | QA_FAIL | 1/8/2/0 |
| 14 | Christian Bedding Set – Inspirational Bible Verse Comforte - Design 05 | 57.6 | QA_FAIL | 1/7/2/0 |
| 15 | Christian Bedding Set – Inspirational Bible Verse Comforte - Design 06 | 45.1 | QA_FAIL | 1/7/2/0 |
| 16 | Christian Bedding Set – Inspirational Bible Verse Comforter | 46.6 | QA_FAIL | 1/6/3/0 |
| 17 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 01 | 85.5 | QA_REVISE | 0/2/1/0 |
| 18 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 02 | 83.5 | QA_REVISE | 0/3/1/0 |
| 19 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 03 | 68.5 | QA_FAIL | 1/3/1/0 |
| 20 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 04 | 74.9 | QA_REVISE | 0/2/2/0 |

## Lỗi ưu tiên

1. **CRITICAL — products 11–16:** draft dùng claim `Personalized`, nhưng purchase options live không có input tên/verse/birth flower; phải chứng minh control hoạt động hoặc bỏ claim.
2. **CRITICAL — product 19:** description ghi `Customization: 1 text input`, nhưng trang live chỉ có size và pillowcase quantity.
3. **MAJOR — cả 10 sản phẩm:** description HTML chứa câu nội bộ yêu cầu review, chưa publish-ready.
4. **MAJOR — ảnh:** alt/observation theo mẫu vị trí sai hàng loạt; nhóm Christian sai nhiều ảnh 3–8, nhóm Christmas thường đảo ảnh cận cảnh với ảnh sham.
5. **MAJOR — products 11–16:** SERP evidence đã nộp chủ yếu là blanket trong khi keyword nhắm comforter/bedding, đồng thời sáu trang cạnh tranh intent rất gần nhau.
6. **MAJOR — product 13:** draft chép `Style: Christian Knight Templar`, mâu thuẫn với thiết kế God Is Within Her dành cho nữ đang hiển thị.

## SERP và keyword

Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword và phương án gần nhất. URL/timestamp/locale limit nằm trong `serp_evidence.json`; không có claim volume.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; chỉ xác minh metadata và media alt ở storefront/product.js, không suy ra giá trị admin.
- Product identity, canonical, product ID, variants và gallery không phát hiện thay đổi trọng yếu so với snapshot; HTML động có thể đổi hash.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã được mở lại, kiểm tra cấu trúc/công thức và render kiểm tra.
- Excel Desktop không thể dùng để recalculate do Office báo `Product Activation Failed`; công thức được kiểm tra cấu trúc, không có `#REF!/#NAME?`, và cả 5 sheet được render độc lập bằng openpyxl/Pillow.
- Chưa QA products 21–30. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_091545\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_091545\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_091545\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_091545`
