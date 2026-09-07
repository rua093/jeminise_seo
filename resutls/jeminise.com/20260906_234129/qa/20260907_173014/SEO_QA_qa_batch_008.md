# SEO QA — qa_batch_008

## Kết luận

- Phạm vi: **10 sản phẩm, 69/69 ảnh (100%)**; chỉ inventory position 71–80.
- Điểm lô: **79.7/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 5 QA_FAIL, 5 QA_REVISE, 0 QA_PASS.
- Phát hiện: 5 CRITICAL, 20 MAJOR, 20 MINOR, 3 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_008.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `1D2CD9D256CCCF553DB09B55AF5463C1636C5CD44B0D2C32E8B530014795151F`

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 71 | Custom Football Player Running Sideways Comforter | 79.7 | QA_REVISE | 0/2/2/0 |
| 72 | Custom Football Player Running with Ball Comforter | 79.7 | QA_REVISE | 0/2/2/0 |
| 73 | Custom Football Player with American Flag Background Comforter | 79.7 | QA_REVISE | 0/2/2/0 |
| 74 | Custom Football Player with Burning Ball Comforter and American Football | 79.7 | QA_REVISE | 0/2/2/0 |
| 75 | Custom Football Players and Large Football Close Comforter | 79.7 | QA_FAIL | 1/2/2/0 |
| 76 | Custom Football Texture with Prominent Laces Vintage Comforter | 79.7 | QA_REVISE | 0/2/2/0 |
| 77 | Custom Glowing Soccer Ball Trails Comforter Set, Machine Washable - A6 | 79.7 | QA_FAIL | 1/2/2/0 |
| 78 | Custom Inspirational Bible Verse Floral Blanket - D5 | 79.8 | QA_FAIL | 1/2/2/0 |
| 79 | Custom Inspirational Floral Butterfly Bible Verse Blanket - D4 | 79.8 | QA_FAIL | 1/2/2/0 |
| 80 | Custom Inspirational Scripture Butterflies Blanket, Machine Washable | 79.8 | QA_FAIL | 1/2/2/0 |

## Lỗi ưu tiên

1. **CRITICAL — products 75, 77–80:** static live Customizer không expose đầy đủ name/number schema như claim; cần kiểm tra control và fulfillment mapping hoặc bỏ claim personalized/custom.
2. **MAJOR — cả 10 sản phẩm:** description HTML chứa khối SEO Use/QA/import nội bộ, chưa publish-ready.
3. **MAJOR — products 71–76:** các landing page football có intent personalized comforter rất gần nhau; cần phân vai keyword và internal linking.
4. **MINOR — ảnh:** alt/observation còn dùng nhãn vị trí chung thay vì mô tả nội dung cụ thể.

## SERP và keyword

Mỗi sản phẩm đã được kiểm tra bằng primary keyword và comparator gần nhất theo US commercial intent; URL, timestamp và locale limit nằm trong `serp_evidence.json`; không có claim volume.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; không suy ra giá trị admin SEO hoặc alt hiện tại.
- Trình duyệt tương tác không khả dụng; static live HTML/product.js đã đọc nhưng chưa chạy Customize-to-cart end-to-end.
- Product identity, canonical, product ID, variants và gallery được đối chiếu với snapshot; không quy kết SOURCE_CHANGED khi chưa có revision evidence.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA gồm đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink.
- Chưa QA products 81–90. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_173014\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_173014\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_173014\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_173014`
