# SEO Re-QA — qa_batch_004_r2

## Kết luận

- Phạm vi: **10 sản phẩm, 62/62 ảnh (100%)**; chỉ inventory position 31–40; revision **r2**.
- Điểm lô: **72.6/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 6 QA_FAIL, 4 QA_REVISE, 0 QA_PASS.
- Phát hiện: 0 CRITICAL, 27 MAJOR, 23 MINOR, 1 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\revisions\qa_batch_004_r2\SEO_Product_Optimization_qa_batch_004_r2.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `EAD02B9EBC6899A9750E8BB1F36C2BC176335A951EFCF5B6844D6566A00730B3`

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 31 | Christmas Cardinals Near Snowy Birdhouse Patchwork Floral Quilt | 83.9 | QA_REVISE | 0/3/3/0 |
| 32 | Christmas Cardinals on Snowy Branches Patchwork Winter Quilt | 84.0 | QA_REVISE | 0/3/2/0 |
| 33 | Cow with Landscape Patchwork Quilt, Animal Farmhouse Printed Quilt Set | 88.9 | QA_REVISE | 0/2/3/0 |
| 34 | Crocodile Patchwork Printed Quilts, Animal Patchwork Bedding | 78.9 | QA_REVISE | 0/3/3/0 |
| 35 | Custom American Football Close Up on Flag Comforter | 68.3 | QA_FAIL | 0/2/2/0 |
| 36 | Custom American Football Close Up on Grunge Vintage Comforter | 60.8 | QA_FAIL | 0/3/2/0 |
| 37 | Custom American Football on Cosmic Background Vintage Comforter | 63.3 | QA_FAIL | 0/3/2/0 |
| 38 | Custom American Football on Usa Flag Background Comforter | 65.8 | QA_FAIL | 0/3/2/0 |
| 39 | Custom American Football with Paint Splash Background Vintage Comforter | 65.8 | QA_FAIL | 0/2/2/0 |
| 40 | Custom American Football with Patriotic Flag Comforter | 65.8 | QA_FAIL | 0/3/2/0 |

## Đối chiếu revision r2

- Đã bỏ đúng khối SEO/QA nội bộ, thuộc tính `Dragonfly` ở products 33–34 và từ khóa `alligator` không có bằng chứng ở product 34.
- Customizer live xác nhận products 35–40 có `Enter Name` bắt buộc (1–25 ký tự) và `Enter Number` tùy chọn (1–5 ký tự); vì vậy các CRITICAL personalization của r1 được gỡ.
- Tuy nhiên r2 bỏ luôn personalization khỏi keyword/copy của products 35–40, làm mất một thuộc tính mua hàng đã được xác minh.
- 5 ảnh vẫn bị gán sai loại và 23 ảnh còn observation/alt theo mẫu chung.

## Lỗi ưu tiên

1. **MAJOR — products 35–40:** phải đưa personalization đã xác minh trở lại keyword/copy và nói rõ tên bắt buộc, số tùy chọn; tên/số trong ảnh chỉ là mẫu.
2. **MAJOR — cả 10 sản phẩm:** description đã sạch câu nội bộ nhưng vẫn quá generic, thiếu thông số, thành phần, care và điều kiện mua quan trọng.
3. **MAJOR — ảnh:** products 31, 33, 34 vẫn gọi image 6 là size guide; product 32 vẫn đảo close-up và angled view.
4. **MAJOR — keywords:** product 34 có evidence exact yếu; product 36 dùng cụm không tự nhiên; product 37 SERP lệch intent; products 38 và 40 trùng intent.
5. **MINOR — 23 ảnh:** alt/observation vẫn dùng nhãn vị trí chung thay vì nội dung trực tiếp của ảnh.

## SERP và keyword

Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword r2 và một phương án gần nhất theo US intent. URL/timestamp/locale limit nằm trong `serp_evidence.json`; không có claim volume.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; chỉ xác minh metadata và media alt ở storefront/product.js, không suy ra giá trị admin.
- Product identity, canonical, product ID, variants và gallery không phát hiện thay đổi trọng yếu so với snapshot; HTML động có thể đổi hash.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại và render kiểm tra.
- Chưa QA products 41–50. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_144539\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_144539\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_144539\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_144539`
