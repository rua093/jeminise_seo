# SEO QA — qa_batch_005

## Kết luận

- Phạm vi: **10 sản phẩm, 72/72 ảnh (100%)**; chỉ inventory position 41–50.
- Điểm lô: **58.2/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 7 QA_FAIL, 3 QA_REVISE, 0 QA_PASS.
- Phát hiện: 7 CRITICAL, 58 MAJOR, 4 MINOR, 2 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_034.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `609EF38D46896F234EA7AAEEC74F393267003D20D6DE8C5BA9BD21BFB19675EE`

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 41 | Custom American Sports Football Player Running with Ball Comforter | 50.8 | QA_FAIL | 1/2/0/0 |
| 42 | Custom Basketball Players Blanket with Name and Number | 45.0 | QA_FAIL | 1/8/0/0 |
| 43 | Custom Bible Verse Photo Blanket – Personalized Christian Gift for Her | 46.8 | QA_FAIL | 1/5/1/0 |
| 44 | Custom Cardinals on Flowering Branches Quilt with Name | 74.5 | QA_REVISE | 0/6/1/0 |
| 45 | Custom Celtic Tree of Life Quilt Set with Colorful Mosaic Sunburst | 78.2 | QA_REVISE | 0/5/1/0 |
| 46 | Custom Celtic Tree of Life Quilt Set with Fantasy Central Eye Artwork | 80.7 | QA_REVISE | 0/5/1/0 |
| 47 | Custom Celtic Tree of Life Quilt Set with Intertwined Roots Design | 70.7 | QA_FAIL | 1/5/0/0 |
| 48 | Custom Christian Faith Semi Truck in Front of Large Comforter | 46.8 | QA_FAIL | 1/5/0/0 |
| 49 | Custom Christian girl Bible emergency numbers Blanket - D12 | 43.3 | QA_FAIL | 1/8/0/0 |
| 50 | Custom Christian Inspirational Sophia Blanket, Machine Washable - D11 | 44.7 | QA_FAIL | 1/9/0/0 |

## Lỗi ưu tiên

1. **CRITICAL — products 41–43, 47–50:** draft dùng personalization/customization nhưng purchase options live không có input name, number, text hoặc photo; phải chứng minh purchase flow và fulfillment mapping hoặc bỏ claim.
2. **MAJOR — cả 10 sản phẩm:** description HTML chứa khối `SEO Use` và câu QA/import nội bộ, chưa publish-ready.
3. **MAJOR — ảnh:** nhiều observation/alt vẫn là câu mẫu theo vị trí; nhóm basketball, Bible-photo và hai Christian blanket sai hàng loạt.
4. **MAJOR — product 42:** gallery đổi mẫu personalization từ COLÓN 06 sang RASHAD 22 ở image 7 nhưng draft không phân biệt.
5. **MAJOR — products 49–50:** draft chép design ID D14 dù handle/title lần lượt là D12 và D11.
6. **MAJOR — products 44–46:** meta description vẫn yêu cầu xem personalization options dù live chỉ có size/pillowcase.

## SERP và keyword

Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword và phương án gần nhất. URL/timestamp/locale limit nằm trong `serp_evidence.json`; không có claim volume. Các trang cạnh tranh cho intent personalized thường có trường nhập tên/số/ảnh rõ ràng, làm nổi bật khoảng trống bằng chứng của purchase flow hiện tại.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; chỉ xác minh metadata và media alt ở storefront/product.js, không suy ra giá trị admin.
- HTML snapshot của giai đoạn 1 bị HTTP 403 nên không thể so trực tiếp metadata cũ; product ID, title, body HTML, options và gallery URL/order trong product JSON cũ khớp nguồn live. Trạng thái revision là `LIVE_PRODUCT_JSON_EQUIVALENT_SNAPSHOT_HTML_UNAVAILABLE`, không quy kết `SOURCE_CHANGED`.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại và render kiểm tra.
- Chưa QA products 51–60. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_101050\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_101050\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_101050\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_101050`
