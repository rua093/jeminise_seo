# SEO QA — qa_batch_006

## Kết luận

- Phạm vi: **10 sản phẩm, 78/78 ảnh (100%)**; chỉ inventory position 51–60.
- Điểm lô: **52.0/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 10 QA_FAIL, 0 QA_REVISE, 0 QA_PASS.
- Phát hiện: 10 CRITICAL, 71 MAJOR, 3 MINOR, 2 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_034.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `609EF38D46896F234EA7AAEEC74F393267003D20D6DE8C5BA9BD21BFB19675EE`

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 51 | Custom Christian scripture floral Blanket, Machine Washable - D10 | 49.3 | QA_FAIL | 1/9/0/0 |
| 52 | Custom Emily God Says I Am Blanket, Machine Washable - D9 | 50.0 | QA_FAIL | 1/9/0/0 |
| 53 | Custom Exploding Soccer Ball Comforter Set, Machine Washable - A10 | 51.2 | QA_FAIL | 1/7/0/0 |
| 54 | Custom Fiery Soccer Ball, Custom Name Comforter Set, Machin - Design 5 | 49.3 | QA_FAIL | 1/8/1/0 |
| 55 | Custom Flaming Soccer Ball with Custom Name Comforter Set - A12 | 50.0 | QA_FAIL | 1/7/0/0 |
| 56 | Custom Flaming Soccer Ball with Name Comforter Set, Machine Washable | 47.5 | QA_FAIL | 1/9/0/0 |
| 57 | Custom Floral Butterfly Inspirational Blanket, Machine Wash - Design 8 | 57.5 | QA_FAIL | 1/4/1/0 |
| 58 | Custom Floral Butterfly Inspirational Blanket, Machine Washable - D3 | 55.0 | QA_FAIL | 1/6/0/0 |
| 59 | Custom Floral Cross Bible Verse Blanket, Machine Washable - D13 | 52.5 | QA_FAIL | 1/8/0/0 |
| 60 | Custom Floral Cross with Butterflies Blanket, Machine Washable - D6 | 57.5 | QA_FAIL | 1/4/1/0 |

## Lỗi ưu tiên

1. **CRITICAL — cả 10 sản phẩm:** draft quảng bá personalization/customization nhưng live purchase options không có input tên, số, ngày hoặc verse; phải chứng minh purchase flow và fulfillment mapping hoặc bỏ claim.
2. **MAJOR — cả 10 sản phẩm:** description HTML chứa câu QA/import nội bộ, chưa publish-ready.
3. **MAJOR — specification:** body ghi polycotton/cotton-polyester, trong khi variants/gallery ghi Fleece/Sherpa hoặc 100% microfiber; cần phân giải vật liệu theo fulfillment variant.
4. **MAJOR — ảnh:** 31/78 observation/alt theo mẫu vị trí không khớp ảnh gốc, đặc biệt các ảnh size, care, material và lifestyle.
5. **MAJOR — products 51, 52, 57–60:** mô tả blanket bị chèn sai cụm `personalized sports bedding`.
6. **MAJOR — products 51–54, 57–60:** design ID chép trong body mâu thuẫn handle/title.
7. **MAJOR — clusters 51–52 và 53–56:** primary intent quá gần nhau, cần phân vai keyword theo thiết kế.

## SERP và keyword

Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword và phương án gần nhất. URL, timestamp và locale limit nằm trong `serp_evidence.json`; không có claim volume. Các trang cạnh tranh nhắm intent personalized đều hiển thị trường nhập dữ liệu rõ ràng, làm nổi bật khoảng trống purchase flow hiện tại.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; chỉ xác minh metadata và media alt ở storefront/product.js, không suy ra giá trị admin.
- HTML snapshot giai đoạn 1 bị HTTP 403; product ID, title, body HTML, options và gallery URL/order trong product JSON cũ khớp live. Trạng thái revision là `LIVE_PRODUCT_JSON_EQUIVALENT_SNAPSHOT_HTML_UNAVAILABLE`, không quy kết `SOURCE_CHANGED`.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại và render kiểm tra.
- Chưa QA products 61–70. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_102634\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_102634\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_102634\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_102634`
