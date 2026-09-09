# SEO Re-QA A-Z độc lập — qa_batch_020_r5

## Kết luận

- Phạm vi: **10 sản phẩm, 46/46 ảnh (100%)**; chỉ inventory position **191–200**; revision **r5**.
- Điểm lô: **88.2/100**; kết luận lô: **PASSED**.
- Trạng thái: 0 QA_FAIL, 0 QA_REVISE, 10 QA_PASS, 0 QA_INCOMPLETE.
- Phát hiện: 0 CRITICAL, 0 MAJOR, 0 MINOR, 1 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\revisions\qa_batch_020_r5\SEO_Product_Optimization_qa_batch_020_r5.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `FA070F63BBF3D2D9F6304C1D4FC20F4AB6723ED22AE581A98DC1531621178C31`
- Ghi chú: batch 20 được chấm độc lập từ r5; QA r3 cũ chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 191 | Custom Red Black Batter Bedding | 87.5 | QA_PASS | 0/0/0/0 |
| 192 | Custom Brown Baseball Glove Bedding | 92.5 | QA_PASS | 0/0/0/0 |
| 193 | Custom Flaming Baseball Batter Bedding | 87.5 | QA_PASS | 0/0/0/0 |
| 194 | Custom Orange Fire Baseball Bedding | 85.0 | QA_PASS | 0/0/0/0 |
| 195 | Custom Baseball Glove Flag Bedding | 92.5 | QA_PASS | 0/0/0/0 |
| 196 | Custom Black Baseball Glove Bedding | 92.5 | QA_PASS | 0/0/0/0 |
| 197 | Custom Lightning Baseball Bedding | 85.0 | QA_PASS | 0/0/0/0 |
| 198 | Custom Night Field Baseball Bedding | 85.0 | QA_PASS | 0/0/0/0 |
| 199 | Custom Smoke Baseball Batter Bedding | 87.5 | QA_PASS | 0/0/0/0 |
| 200 | Custom Baseball Stitch Name Bedding | 87.5 | QA_PASS | 0/0/0/0 |

## Lỗi và giới hạn ưu tiên

1. **LIMITATION — revision traceability** (batch-level): Traceability limitation only; it does not prove the submitted r5 content is wrong. Đề xuất: Update row-level revision values and revision_summary paths in the next package.

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Cụm baseball bedding/custom-name có intent mua hàng rõ; các exact modifier như red-black batter, orange fire, lightning, night field, smoke và stitch được chấm thận trọng khi SERP còn rộng hoặc lẫn blanket/duvet/decor ngoài bedding.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 46 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r5 khớp các motif chính: red/black batter, brown/black glove, flaming/orange fire, glove flag, lightning, night field, smoke batter và baseball stitch/name artwork.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 46 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
- Test logic đạt: `100 + CRITICAL -> QA_FAIL`, `90 + đủ coverage + không lỗi chặn -> QA_PASS`, `72/80 -> 72–92 và QA_INCOMPLETE`.
- XLSX có đúng 5 sheet: `QA_Summary`, `QA_Products`, `QA_Criteria`, `QA_Images`, `QA_Issues`; có filter/freeze/wrap/hyperlink và công thức truy kiểm.
- Không tạo `rendered_sheets` trong project; preview nằm trong thư mục tạm ngoài project.
- Không sửa workbook nguồn, không tạo APPROVED/import, không cập nhật Shopify. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA workbook: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\qa\20260909_131100\SEO_QA_qa_batch_020_r5.xlsx`
- QA report: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\qa\20260909_131100\SEO_QA_qa_batch_020_r5.md`
- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260909_131100\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260909_131100\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260909_131100\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260909_131100`
