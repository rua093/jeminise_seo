# SEO QA — qa_batch_004

## Kết luận

- Phạm vi: **10 sản phẩm, 62/62 ảnh (100%)**; chỉ inventory position 31–40.
- Điểm lô: **65.2/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 8 QA_FAIL, 2 QA_REVISE, 0 QA_PASS.
- Phát hiện: 8 CRITICAL, 21 MAJOR, 4 MINOR, 1 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_034.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `609EF38D46896F234EA7AAEEC74F393267003D20D6DE8C5BA9BD21BFB19675EE`

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 31 | Christmas Cardinals Near Snowy Birdhouse Patchwork Floral Quilt | 81.1 | QA_REVISE | 0/4/1/0 |
| 32 | Christmas Cardinals on Snowy Branches Patchwork Winter Quilt | 81.0 | QA_REVISE | 0/4/1/0 |
| 33 | Cow with Landscape Patchwork Quilt, Animal Farmhouse Printed Quilt Set | 73.6 | QA_FAIL | 1/3/1/0 |
| 34 | Crocodile Patchwork Printed Quilts, Animal Patchwork Bedding | 68.6 | QA_FAIL | 1/4/1/0 |
| 35 | Custom American Football Close Up on Flag Comforter | 57.5 | QA_FAIL | 1/1/0/0 |
| 36 | Custom American Football Close Up on Grunge Vintage Comforter | 57.5 | QA_FAIL | 1/1/0/0 |
| 37 | Custom American Football on Cosmic Background Vintage Comforter | 60.0 | QA_FAIL | 1/1/0/0 |
| 38 | Custom American Football on Usa Flag Background Comforter | 57.5 | QA_FAIL | 1/1/0/0 |
| 39 | Custom American Football with Paint Splash Background Vintage Comforter | 57.5 | QA_FAIL | 1/1/0/0 |
| 40 | Custom American Football with Patriotic Flag Comforter | 57.5 | QA_FAIL | 1/1/0/0 |

## Lỗi ưu tiên

1. **CRITICAL — products 33–40:** draft dùng customization/personalized/custom nhưng purchase options live không có input name/number/text; phải chứng minh control và fulfillment mapping hoặc bỏ claim.
2. **MAJOR — cả 10 sản phẩm:** description HTML chứa khối `SEO Use` và câu QA/import nội bộ, chưa publish-ready.
3. **MAJOR — products 33–34:** draft chép thuộc tính `Dragonfly`, mâu thuẫn với thiết kế cow/crocodile.
4. **MAJOR — ảnh:** products 31, 33, 34 gán image 6 là size guide thay vì material-layer diagram; product 32 đảo close-up và angled bedroom.
5. **MAJOR — product 31:** SERP exact birdhouse còn lẫn fabric panel/quilt kit; finished-product intent chưa được xác nhận sạch.
6. **MAJOR — product 34:** dùng `alligator` như secondary keyword khi evidence sản phẩm gọi là crocodile.

## SERP và keyword

Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword và phương án gần nhất. URL/timestamp/locale limit nằm trong `serp_evidence.json`; không có claim volume. Các SERP football có sản phẩm tùy biến thật với input name/number, làm rõ khoảng cách giữa intent keyword và purchase flow hiện tại.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; chỉ xác minh metadata và media alt ở storefront/product.js, không suy ra giá trị admin.
- Trình duyệt tương tác không khả dụng trong môi trường; static HTML/product.js không hiển thị input personalization. Claim chỉ được thông qua lại khi có bằng chứng purchase flow hoạt động.
- Product identity, canonical, product ID, variants và gallery không phát hiện thay đổi trọng yếu so với snapshot; HTML động có thể đổi hash.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại và render kiểm tra.
- Chưa QA products 41–50. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_095643\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_095643\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_095643\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_095643`
