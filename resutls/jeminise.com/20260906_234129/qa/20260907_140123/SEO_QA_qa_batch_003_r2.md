# SEO Re-QA — qa_batch_003_r2

## Kết luận

- Phạm vi: **10 sản phẩm, 50/50 ảnh (100%)**; chỉ inventory position 21–30; revision **r2**.
- Điểm lô: **75.2/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 2 QA_FAIL, 8 QA_REVISE, 0 QA_PASS.
- Phát hiện: 2 CRITICAL, 37 MAJOR, 8 MINOR, 1 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\revisions\qa_batch_003_r2\SEO_Product_Optimization_qa_batch_003_r2.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `D7284B208FD4B5DB5B98B0AC5215CD835B7F42CDC858B5F27899F6F0B9CFDC20`

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 21 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 05 | 76.0 | QA_REVISE | 0/5/1/0 |
| 22 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 06 | 87.5 | QA_REVISE | 0/2/1/0 |
| 23 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 08 | 82.5 | QA_REVISE | 0/3/1/0 |
| 24 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 09 | 87.5 | QA_REVISE | 0/2/1/0 |
| 25 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 10 | 31.0 | QA_FAIL | 1/7/1/0 |
| 26 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 11 | 87.5 | QA_REVISE | 0/2/1/0 |
| 27 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 12 | 41.0 | QA_FAIL | 1/7/1/0 |
| 28 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedspread | 87.5 | QA_REVISE | 0/2/1/0 |
| 29 | Christmas Cardinal on Branch Within Patchwork Winter Quilt | 88.5 | QA_REVISE | 0/3/0/0 |
| 30 | Christmas Cardinal Perched on Branch Within Wreath Patchwork Winter Quilt | 83.5 | QA_REVISE | 0/4/0/0 |

## Đối chiếu revision r2

- Đã sửa đúng product 24 từ snowman thành gingerbread và loại bỏ câu QA/import nội bộ khỏi 10 description.
- Tuy nhiên, r2 tạo lỗi motif mới ở products 25 và 27: ảnh là snowman nhưng keyword/title/meta/alt lại ghi gingerbread.
- Sáu alt/observation được báo sửa ở products 21, 29 và 30 vẫn đảo sai close-up, sham hoặc angled view.
- Live Customizer xác nhận trường `Custom Your Name` tùy chọn ở products 21–28; r2 chưa mô tả điều kiện này.

## Lỗi ưu tiên

1. **CRITICAL — products 25, 27:** keyword/title/meta/alt nhận diện snowman thành gingerbread; phải sửa toàn bộ chuỗi keyword → copy → alt và chạy lại SERP decision.
2. **MAJOR — 16 ảnh:** 6 ảnh vẫn đảo sai loại ảnh và 10 alt ở products 25/27 sai motif.
3. **MAJOR — cả 10 sản phẩm:** description đã bỏ câu nội bộ nhưng vẫn quá generic, thiếu kích thước, thành phần quilt/sham và điều kiện mua đã xác minh.
4. **MAJOR — products 21–28:** live Customizer có trường tên tùy chọn nhưng r2 bỏ qua tính năng đã xác minh này.
5. **MAJOR — products 21, 23:** SERP exact-match cho finished quilt set vẫn chưa đủ sạch so với pattern/kit hoặc motif rộng hơn.
6. **MINOR — products 21–28:** title/H1 live còn ký tự `�` và từ `Bedsprea` bị cắt.

## SERP và keyword

Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword r2 và một phương án gần nhất. URL/timestamp/locale limit nằm trong `serp_evidence.json`; không có claim volume.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; chỉ xác minh metadata và media alt ở storefront/product.js, không suy ra giá trị admin.
- Product identity, canonical, product ID, variants và gallery không phát hiện thay đổi trọng yếu so với snapshot; HTML động có thể đổi hash.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã được mở lại và render kiểm tra.
- Chưa QA products 31–40 trong lượt này. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_140123\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_140123\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_140123\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_140123`
