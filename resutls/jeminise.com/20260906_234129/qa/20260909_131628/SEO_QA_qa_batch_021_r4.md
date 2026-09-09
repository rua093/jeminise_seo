# SEO Re-QA A-Z độc lập — qa_batch_021_r4

## Kết luận

- Phạm vi: **10 sản phẩm, 64/64 ảnh (100%)**; chỉ inventory position **201–210**; revision **r4**.
- Điểm lô: **94.8/100**; kết luận lô: **PASSED**.
- Trạng thái: 0 QA_FAIL, 0 QA_REVISE, 10 QA_PASS, 0 QA_INCOMPLETE.
- Phát hiện: 0 CRITICAL, 0 MAJOR, 0 MINOR, 1 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\revisions\qa_batch_021_r4\SEO_Product_Optimization_qa_batch_021_r4.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `47771C12438547A1A05FCA12AA6D3FD79345B043DA005855C1B45A903F9004AC`
- Ghi chú: batch 21 được chấm độc lập từ r4; QA r2/r3 cũ chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 201 | Custom Baseball Flag Name Number Bedding | 97.5 | QA_PASS | 0/0/0/0 |
| 202 | Custom Neon Baseball Player Duvet Cover | 90.0 | QA_PASS | 0/0/0/0 |
| 203 | Custom Basketball Hoop Comforter Set | 97.5 | QA_PASS | 0/0/0/0 |
| 204 | Custom Blue Red Basketball Hoop Bedding | 92.5 | QA_PASS | 0/0/0/0 |
| 205 | Custom Basketball Paint Splash Comforter | 97.5 | QA_PASS | 0/0/0/0 |
| 206 | Custom Basketball Net Blanket | 97.5 | QA_PASS | 0/0/0/0 |
| 207 | Custom Basketball Court Hoop Comforter | 97.5 | QA_PASS | 0/0/0/0 |
| 208 | Custom Black Basketball Hoop Comforter | 92.5 | QA_PASS | 0/0/0/0 |
| 209 | Custom Flame Basketball Comforter Set | 92.5 | QA_PASS | 0/0/0/0 |
| 210 | Custom Basketball Close-Up Blanket | 92.5 | QA_PASS | 0/0/0/0 |

## Lỗi và giới hạn ưu tiên

1. **LIMITATION — revision traceability** (batch-level): Traceability limitation only; it does not prove the submitted r4 content is wrong. Đề xuất: Update row-level revision values and revision_summary paths in the next package.

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Baseball và basketball custom-name/name-number có intent mua hàng rõ; các exact modifier như neon baseball player, blue-red hoop, black hoop, flame và close-up được chấm thận trọng khi SERP còn rộng hoặc marketplace-heavy.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 64 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r4 khớp các motif chính: baseball flag/name-number, neon baseball player, basketball hoop/net/court, paint splash, flame, close-up ball và blanket/comforter product forms.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 64 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
- Test logic đạt: `100 + CRITICAL -> QA_FAIL`, `90 + đủ coverage + không lỗi chặn -> QA_PASS`, `72/80 -> 72–92 và QA_INCOMPLETE`.
- XLSX có đúng 5 sheet: `QA_Summary`, `QA_Products`, `QA_Criteria`, `QA_Images`, `QA_Issues`; có filter/freeze/wrap/hyperlink và công thức truy kiểm.
- Không tạo `rendered_sheets` trong project; preview nằm trong thư mục tạm ngoài project.
- Không sửa workbook nguồn, không tạo APPROVED/import, không cập nhật Shopify. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA workbook: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\qa\20260909_131628\SEO_QA_qa_batch_021_r4.xlsx`
- QA report: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\qa\20260909_131628\SEO_QA_qa_batch_021_r4.md`
- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260909_131628\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260909_131628\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260909_131628\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260909_131628`
