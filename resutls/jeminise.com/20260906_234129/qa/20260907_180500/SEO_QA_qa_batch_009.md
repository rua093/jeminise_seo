# SEO QA — qa_batch_009

## Kết luận

- Phạm vi: **10 sản phẩm, 59/59 ảnh (100%)**; chỉ inventory position 81–90.
- Điểm lô: **79.7/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 5 QA_FAIL, 5 QA_REVISE, 0 QA_PASS.
- Phát hiện: 5 CRITICAL, 20 MAJOR, 20 MINOR, 3 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_009.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `D1D265F3AD9921852CF46DFC086093B832ABE511E03DA9F0E50B5174925A26EA`

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 81 | Custom Name Catcher Batter Baseball Baseball Throw, Baseba - Design 01 | 79.6 | QA_REVISE | 0/2/2/0 |
| 82 | Custom Name Catcher Batter Baseball Baseball Throw, Baseba - Design 07 | 79.7 | QA_REVISE | 0/2/2/0 |
| 83 | Custom Name Catcher Batter Baseball Baseball Throw, Baseba - Design 10 | 79.7 | QA_REVISE | 0/2/2/0 |
| 84 | Custom Name Catcher Batter Baseball Baseball Throw, Baseba - Design 11 | 79.7 | QA_REVISE | 0/2/2/0 |
| 85 | Custom Name Catcher Batter Baseball Baseball Throw, Baseba - Design 12 | 79.7 | QA_FAIL | 1/2/2/0 |
| 86 | Custom Name Catcher Batter Baseball Baseball Throw, Baseba - Design 13 | 79.7 | QA_REVISE | 0/2/2/0 |
| 87 | Custom Name Catcher Batter Baseball Baseball Throw, Baseba - Design 17 | 79.7 | QA_FAIL | 1/2/2/0 |
| 88 | Custom Name Catcher Batter Baseball Baseball Throw, Baseba - Design 18 | 79.7 | QA_FAIL | 1/2/2/0 |
| 89 | Custom Name Catcher Batter Baseball Baseball Throw, Baseba - Design 21 | 79.6 | QA_FAIL | 1/2/2/0 |
| 90 | Custom Name Catcher Batter Baseball Baseball Throw, Baseball Blanket | 79.5 | QA_FAIL | 1/2/2/0 |

## Lỗi ưu tiên

1. **CRITICAL — products 85, 87–90:** static live Customizer không expose đầy đủ name/number schema như claim; cần kiểm tra control và fulfillment mapping hoặc bỏ claim personalized/custom.
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

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_180500\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_180500\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_180500\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_180500`
