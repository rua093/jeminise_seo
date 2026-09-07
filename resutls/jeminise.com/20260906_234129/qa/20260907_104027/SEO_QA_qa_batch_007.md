# SEO QA — qa_batch_007

## Kết luận

- Phạm vi: **10 sản phẩm, 60/60 ảnh (100%)**; chỉ inventory position 61–70.
- Điểm lô: **79.4/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 0 QA_FAIL, 10 QA_REVISE, 0 QA_PASS.
- Phát hiện: 0 CRITICAL, 22 MAJOR, 19 MINOR, 3 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_034.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `609EF38D46896F234EA7AAEEC74F393267003D20D6DE8C5BA9BD21BFB19675EE`

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 61 | Custom Football Above Yard Lines Perspective Comforter | 79.7 | QA_REVISE | 0/2/2/0 |
| 62 | Custom Football Close Up on American Flag Comforter | 79.7 | QA_REVISE | 0/2/2/0 |
| 63 | Custom Football Engulfed in Flames and Lightning Comforter | 79.7 | QA_REVISE | 0/2/2/0 |
| 64 | Custom Football Helmet on American Flag Vintage Comforter | 79.7 | QA_REVISE | 0/2/2/0 |
| 65 | Custom Football Helmet with Flaming Football Comforter | 76.7 | QA_REVISE | 0/4/1/0 |
| 66 | Custom Football Over American Flag Background Comforter | 79.7 | QA_REVISE | 0/2/2/0 |
| 67 | Custom Football Player Holding Ball Close Up Comforter | 79.7 | QA_REVISE | 0/2/2/0 |
| 68 | Custom Football Player Holding Ball Close Up Comforter with Football Helmet | 79.7 | QA_REVISE | 0/2/2/0 |
| 69 | Custom Football Player Holding Ball Collage Comforter with Football Helmet | 79.7 | QA_REVISE | 0/2/2/0 |
| 70 | Custom Football Player Running on Camouflage Flag Comforter | 79.7 | QA_REVISE | 0/2/2/0 |

## Lỗi ưu tiên

1. **MAJOR — cả 10 sản phẩm:** description HTML chứa khối `SEO Use` và câu QA/import nội bộ, chưa publish-ready.
2. **MAJOR — cluster football:** 10 trang nhắm intent personalized football bedding rất gần nhau; cần phân vai keyword theo yard-line, patriotic, flame, helmet, player và camo.
3. **MAJOR — product 65, ảnh 4–5:** observation/alt bị đảo giữa easy-care panel và feature panel.
4. **MINOR — cả 10 sản phẩm:** meta description dài 169–200 ký tự; cần rút gọn và nói rõ người mua nhập name/number qua Customize.
5. **MINOR — 9 ảnh feature panel:** alt chỉ gọi là close-up, chưa nêu đúng mục đích bảng soft/lightweight/durable/breathable.

## Personalization

- Không phát hiện lỗi chặn personalization trong batch này. Cả 10 HTML live có Amazon Customizer, nút Customize, endpoint upload, trường Enter Name bắt buộc (1–25 ký tự) và Enter Number tùy chọn (1–5 ký tự).
- Chưa chạy được thao tác tương tác đến cart/fulfillment payload; đây là LIMITATION, không phải bằng chứng claim sai.

## SERP và keyword

Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword và phương án gần nhất. URL, timestamp và locale limit nằm trong `serp_evidence.json`; không có claim volume. SERP cho thấy personalized football bedding là intent thương mại có thật, nhưng nhiều kết quả là blanket hoặc duvet nên evidence được chấm PARTIAL.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; chỉ xác minh metadata, media alt và customizer schema ở storefront/product.js/HTML live, không suy ra giá trị admin.
- HTML snapshot giai đoạn 1 bị HTTP 403; product ID, title, body HTML, options và gallery URL/order trong product JSON cũ khớp live. Trạng thái revision là `LIVE_PRODUCT_JSON_EQUIVALENT_SNAPSHOT_HTML_UNAVAILABLE`, không quy kết `SOURCE_CHANGED`.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại và render kiểm tra.
- Chưa QA products 71–80. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_104027\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_104027\serp_evidence.json`
- Customizer audit: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_104027\customizer_audit.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_104027\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_104027`
