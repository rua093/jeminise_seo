# SEO QA — qa_batch_003

## Kết luận

- Phạm vi: **10 sản phẩm, 50/50 ảnh (100%)**; chỉ inventory position 21–30.
- Điểm lô: **77.3/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 4 QA_FAIL, 6 QA_REVISE, 0 QA_PASS.
- Phát hiện: 4 CRITICAL, 18 MAJOR, 8 MINOR, 1 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_034.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `609EF38D46896F234EA7AAEEC74F393267003D20D6DE8C5BA9BD21BFB19675EE`

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 21 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 05 | 66.0 | QA_FAIL | 1/4/1/0 |
| 22 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 06 | 87.5 | QA_REVISE | 0/1/1/0 |
| 23 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 08 | 80.0 | QA_REVISE | 0/2/1/0 |
| 24 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 09 | 42.5 | QA_FAIL | 1/1/1/0 |
| 25 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 10 | 87.5 | QA_REVISE | 0/1/1/0 |
| 26 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 11 | 87.5 | QA_REVISE | 0/1/1/0 |
| 27 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 12 | 87.5 | QA_REVISE | 0/1/1/0 |
| 28 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedspread | 87.5 | QA_REVISE | 0/1/1/0 |
| 29 | Christmas Cardinal on Branch Within Patchwork Winter Quilt | 73.5 | QA_FAIL | 1/3/0/0 |
| 30 | Christmas Cardinal Perched on Branch Within Wreath Patchwork Winter Quilt | 73.5 | QA_FAIL | 1/3/0/0 |

## Lỗi ưu tiên

1. **CRITICAL — products 21, 29, 30:** draft ghi customization/personalization nhưng purchase options live không có text/name input; phải chứng minh control hoạt động hoặc bỏ claim.
2. **CRITICAL — product 24:** draft gọi nhân vật trung tâm là snowman, trong khi ảnh là gingerbread figure; cần thay keyword/title/meta/body và chạy lại SERP decision.
3. **MAJOR — cả 10 sản phẩm:** description HTML chứa câu QA/import nội bộ, chưa publish-ready.
4. **MAJOR — ảnh:** product 21 đảo close-up với sham; products 29–30 đảo close-up với angled bedroom view.
5. **MAJOR — products 21, 23:** SERP primary còn lẫn pattern/kit và chưa sạch intent finished quilt set.
6. **MINOR — products 21–28:** title/H1 nguồn có ký tự `�` và từ `Bedsprea` bị cắt.

## SERP và keyword

Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword và phương án gần nhất. URL/timestamp/locale limit nằm trong `serp_evidence.json`; không có claim volume.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; chỉ xác minh metadata và media alt ở storefront/product.js, không suy ra giá trị admin.
- Product identity, canonical, product ID, variants và gallery không phát hiện thay đổi trọng yếu so với snapshot; HTML động có thể đổi hash.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã được mở lại, kiểm tra cấu trúc/công thức và render kiểm tra.
- Excel Desktop không thể dùng để recalculate do Office báo `Product Activation Failed`; công thức được kiểm tra cấu trúc, không có `#REF!/#NAME?`, và cả 5 sheet được render độc lập bằng openpyxl/Pillow.
- Chưa QA products 31–40. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_094127\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_094127\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_094127\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_094127`
