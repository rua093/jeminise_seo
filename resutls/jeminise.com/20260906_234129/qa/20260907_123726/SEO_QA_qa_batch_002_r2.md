# SEO Re-QA — qa_batch_002_r2

## Kết luận

- Phạm vi: **10 sản phẩm, 64/64 ảnh (100%)**; chỉ inventory position 11–20; revision **r2**.
- Điểm lô: **59.8/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 7 QA_FAIL, 3 QA_REVISE, 0 QA_PASS.
- Phát hiện: 0 CRITICAL, 55 MAJOR, 20 MINOR, 1 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\revisions\qa_batch_002_r2\SEO_Product_Optimization_qa_batch_002_r2.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `75C50FAC073CBFE79A77C7A664BA8006EDD9327790CA0AC2431E196534DC9A7A`

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 11 | Christian Bedding Set – Inspirational Bible Verse Comforte - Design 02 | 46.0 | QA_FAIL | 0/6/4/0 |
| 12 | Christian Bedding Set – Inspirational Bible Verse Comforte - Design 03 | 45.1 | QA_FAIL | 0/7/2/0 |
| 13 | Christian Bedding Set – Inspirational Bible Verse Comforte - Design 04 | 45.1 | QA_FAIL | 0/8/2/0 |
| 14 | Christian Bedding Set – Inspirational Bible Verse Comforte - Design 05 | 57.6 | QA_FAIL | 0/7/2/0 |
| 15 | Christian Bedding Set – Inspirational Bible Verse Comforte - Design 06 | 45.1 | QA_FAIL | 0/7/2/0 |
| 16 | Christian Bedding Set – Inspirational Bible Verse Comforter | 46.6 | QA_FAIL | 0/6/3/0 |
| 17 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 01 | 85.5 | QA_REVISE | 0/3/1/0 |
| 18 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 02 | 83.5 | QA_REVISE | 0/4/1/0 |
| 19 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 03 | 68.5 | QA_FAIL | 0/4/1/0 |
| 20 | Christmas Bedding Quilt – Soft Microfiber Holiday Bedsprea - Design 04 | 74.9 | QA_REVISE | 0/3/2/0 |

## Lỗi ưu tiên

1. **MAJOR — toàn bộ 64 ảnh:** nhiều observation/alt r2 vẫn gán theo vị trí chung hoặc đảo cận cảnh–sham; cần sửa theo nội dung ảnh thực tế.
2. **MAJOR — 10 sản phẩm:** r2 đã bỏ câu nội bộ draft/QA nhưng còn quá generic và chưa nêu đầy đủ dữ kiện đã xác minh.
3. **MAJOR — products 11–20:** live Customizer có trường tên (bắt buộc ở 11–16, tùy chọn ở 17–20), nhưng copy r2 chưa ánh xạ rõ điều kiện personalization.
4. **MAJOR — products 11–16:** SERP còn lệch intent (blanket so với comforter/bedding) và các trang cùng cụm cạnh tranh quá gần nhau.
5. **MAJOR — product 13:** cụm `Christian Knight Templar` vẫn mâu thuẫn với thiết kế God Is Within Her đang hiển thị.
6. **MINOR — một số title/H1 nguồn:** còn ký tự `�` hoặc từ `Bedsprea` bị cắt; cần xác nhận bản live/admin trước khi sửa.

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

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_123726\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_123726\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_123726\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_123726`
