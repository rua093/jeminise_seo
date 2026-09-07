# SEO Re-QA — qa_batch_007_r2

## Kết luận

- Phạm vi: **10 sản phẩm, 60/60 ảnh (100%)**; chỉ inventory position 61–70; revision **r2**.
- Điểm lô: **81.9/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 0 QA_FAIL, 10 QA_REVISE, 0 QA_PASS.
- Phát hiện: 0 CRITICAL, 32 MAJOR, 19 MINOR, 3 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\revisions\qa_batch_007_r2\SEO_Product_Optimization_qa_batch_007_r2.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `5B57B6C549B0A1FF15628193061999DE302D3377FE777B0984F8B0D19B7D7FE9`

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 61 | Custom Football Above Yard Lines Perspective Comforter | 82.2 | QA_REVISE | 0/3/2/0 |
| 62 | Custom Football Close Up on American Flag Comforter | 82.2 | QA_REVISE | 0/3/2/0 |
| 63 | Custom Football Engulfed in Flames and Lightning Comforter | 82.2 | QA_REVISE | 0/3/2/0 |
| 64 | Custom Football Helmet on American Flag Vintage Comforter | 82.2 | QA_REVISE | 0/3/2/0 |
| 65 | Custom Football Helmet with Flaming Football Comforter | 79.2 | QA_REVISE | 0/5/1/0 |
| 66 | Custom Football Over American Flag Background Comforter | 82.2 | QA_REVISE | 0/3/2/0 |
| 67 | Custom Football Player Holding Ball Close Up Comforter | 82.2 | QA_REVISE | 0/3/2/0 |
| 68 | Custom Football Player Holding Ball Close Up Comforter with Football Helmet | 82.2 | QA_REVISE | 0/3/2/0 |
| 69 | Custom Football Player Holding Ball Collage Comforter with Football Helmet | 82.2 | QA_REVISE | 0/3/2/0 |
| 70 | Custom Football Player Running on Camouflage Flag Comforter | 82.2 | QA_REVISE | 0/3/2/0 |

## So sánh với r1

- Điểm lô tăng từ **79.4** lên **81.9** (+2.5 điểm).
- r2 đã bỏ khối SEO/QA/import nội bộ; personalization được xác minh bằng schema Customizer live nên không có CRITICAL.
- Trạng thái vẫn là 10 QA_REVISE, 0 QA_PASS vì mỗi sản phẩm còn ít nhất một lỗi MAJOR.

## Lỗi ưu tiên

1. **MAJOR — 10 sản phẩm:** description r2 vẫn là mẫu chung, lặp `Visible artwork`, thiếu material, care, exact contents và điều kiện option.
2. **MAJOR — 10 sản phẩm:** copy chưa nói rõ Name là bắt buộc (1–25 ký tự), Number là tùy chọn (1–5 ký tự) trong Customizer.
3. **MAJOR — cluster football:** 10 trang vẫn nhắm intent personalized football bedding rất gần nhau; cần phân vai keyword/landing page theo yard-line, patriotic, flame, helmet, player, collage và camo.
4. **MAJOR — product 65, ảnh 4–5:** observation/alt vẫn bị đảo giữa easy-care panel và feature panel.
5. **MINOR — 10 sản phẩm:** meta description dài 168–199 ký tự và dùng CTA chung; cần rút gọn, nêu Customize chính xác.
6. **MINOR — 9 ảnh feature panel:** alt còn gọi chung là close-up thay vì mô tả đúng bảng soft/lightweight/durable/breathable.

## SERP và keyword

Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword và một phương án gần nhất theo US commercial intent. URL, timestamp và locale limit nằm trong `serp_evidence.json`; không có claim volume.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; chỉ xác minh storefront/product.js và Customizer schema, không suy ra giá trị admin.
- Trình duyệt tương tác không khả dụng; đã xác minh schema live nhưng chưa chạy Customize-to-cart để kiểm tra persistence/fulfillment payload.
- Product identity, canonical, product ID, variants và gallery không phát hiện thay đổi trọng yếu so với snapshot; HTML động có thể đổi hash.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại và render kiểm tra.
- Chưa QA products 71–80. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_164400\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_164400\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_164400\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_164400`
