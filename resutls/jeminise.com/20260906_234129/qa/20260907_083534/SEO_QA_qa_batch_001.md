# SEO QA — qa_batch_001

## Kết luận

- Phạm vi: **10 sản phẩm, 65/65 ảnh (100%)**; chỉ inventory position 1–10.
- Điểm lô: **74.3/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 3 QA_FAIL, 7 QA_REVISE, 0 QA_PASS.
- Phát hiện: 2 CRITICAL, 29 MAJOR, 17 MINOR, 3 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_034.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `609EF38D46896F234EA7AAEEC74F393267003D20D6DE8C5BA9BD21BFB19675EE`

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 1 | Autumn Cardinals Perched Above Large Sunflowers Quilt with Cardinal Birds | 76.5 | QA_REVISE | 0/1/1/1 |
| 2 | Bright Teal Yellow Softball Comforter Set – Premium Sports Theme | 77.6 | QA_REVISE | 0/5/1/0 |
| 3 | Cardinals Perched Beside Decorated Christmas Tree Patchwork Winter Quilt | 79.0 | QA_REVISE | 0/2/1/0 |
| 4 | Cardinals Perched on Branch with Roses Patchwork Winter Quilt | 76.5 | QA_REVISE | 0/2/1/0 |
| 5 | Cat Patchwork Quilt Set, Animal Pattern Bedding, Cat Pattern Quilt Set | 70.1 | QA_FAIL | 1/3/2/0 |
| 6 | Cat Quilt Set - Cat Patchwork Style Printed Bedding, Animal Patchwork | 65.1 | QA_FAIL | 0/3/2/1 |
| 7 | Celtic and Fantasy Quilt Set with Twisting Tree Trunk And Exposed Roots | 80.1 | QA_REVISE | 0/2/2/0 |
| 8 | Celtic Mythology Tree of Life Quilt Set With Celtic Knotwork Border | 85.1 | QA_REVISE | 0/2/2/0 |
| 9 | Chicken Pattern Quilt Set, Animal Patchwork Farmhouse Quilt Set | 72.6 | QA_REVISE | 0/3/2/0 |
| 10 | Christian Bedding Set – Inspirational Bible Verse Comforte - Design 01 | 60.8 | QA_FAIL | 1/6/3/0 |

## Lỗi ưu tiên

1. **CRITICAL — product 5:** evidence SERP của cat quilt lại dẫn tới `chicken_quilt`; phải thay nguồn và chạy lại quyết định keyword.
2. **CRITICAL — product 10:** draft dùng claim `Personalized`, nhưng purchase options live không cho thấy input name/birth flower; phải chứng minh control hoạt động hoặc bỏ claim.
3. **MAJOR — cả 10 sản phẩm:** description HTML chứa câu nội bộ về draft/QA và chưa phải nội dung có thể xuất bản.
4. **MAJOR — ảnh:** nhiều observation/alt được gán theo vị trí chung thay vì nội dung thật; product 10 sai nặng ở ảnh 3–8.
5. **MAJOR — products 5, 6, 9:** draft chép `Dragonfly` từ dữ liệu nguồn mâu thuẫn với thiết kế cat/chicken.

## SERP và keyword

Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword và một phương án gần nhất. Các URL/timestamp/locale limit nằm trong `serp_evidence.json`. Không có claim volume. Hai keyword của products 1 và 6 vẫn giữ `HYPOTHESIS_ONLY`.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; chỉ xác minh metadata và media alt ở storefront/product.js, không suy ra giá trị admin.
- Product identity, canonical, product ID, variants, body HTML và gallery filename/order không đổi giữa snapshot và live. HTML trang động có hash khác nhưng không phát hiện drift nội dung sản phẩm trọng yếu.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã được Excel tính lại và render để kiểm tra.
- Chưa QA products 11–20. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_083534\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_083534\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_083534\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_083534`
