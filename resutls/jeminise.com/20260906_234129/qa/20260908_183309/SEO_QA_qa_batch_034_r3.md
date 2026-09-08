# SEO QA — qa_batch_034_r3

## Kết luận

- Phạm vi: **8 sản phẩm, 57/57 ảnh (100%)**; chỉ inventory position 331–338; revision **r3**.
- Điểm lô: **87.5/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 0 QA_FAIL, 8 QA_REVISE, 0 QA_PASS, 0 QA_INCOMPLETE.
- Phát hiện: 0 CRITICAL, 11 MAJOR, 6 MINOR, 8 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\revisions\qa_batch_034_r3\SEO_Product_Optimization_qa_batch_034_r3.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `E8FF901F334297FAD3934DEA63775C95CF1BA9C5DFB1A80F7738B9C370EDC320`
- Ghi chú: meta description r3 đã sửa lỗi cắt cụt; mốc 145–165 ký tự chỉ là biên tập, không bị dùng làm lỗi tự động.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 331 | Vintage Yellow Softball Comforter Set | 85.0 | QA_REVISE | 0/1/2/1 |
| 332 | Polka Dot Softball Comforter Set | 85.0 | QA_REVISE | 0/2/0/1 |
| 333 | Fireball Softball Comforter Set | 85.0 | QA_REVISE | 0/1/2/1 |
| 334 | Pink Glove Softball Comforter Set | 90.0 | QA_REVISE | 0/2/1/1 |
| 335 | Celtic Tree of Life Quilt Set | 92.5 | QA_REVISE | 0/1/0/1 |
| 336 | Winter Cardinal Berry Branch Quilt | 90.0 | QA_REVISE | 0/1/0/1 |
| 337 | Heart Branch Cardinals Christmas Quilt | 85.0 | QA_REVISE | 0/2/0/1 |
| 338 | Winter Cat Cardinal Christmas Quilt | 87.5 | QA_REVISE | 0/1/1/1 |

## Lỗi ưu tiên

1. **MAJOR — product 331:** This is not fully customer-facing publish-ready storefront copy even though the facts are largely verified. Đề xuất: Rewrite in English as customer-facing copy for Vintage Yellow Softball Comforter Set; remove editorial/process wording, keep verified options/customizer labels, and preserve exact motif facts.
2. **MAJOR — product 332:** The selected keyword may be too narrow or cannibalize adjacent products without enough demand proof. Đề xuất: Use a broader primary with the motif as secondary, or add stronger exact SERP/Search Console evidence.
3. **MAJOR — product 332:** This is not fully customer-facing publish-ready storefront copy even though the facts are largely verified. Đề xuất: Rewrite in English as customer-facing copy for Polka Dot Softball Comforter Set; remove editorial/process wording, keep verified options/customizer labels, and preserve exact motif facts.
4. **MAJOR — product 333:** This is not fully customer-facing publish-ready storefront copy even though the facts are largely verified. Đề xuất: Rewrite in English as customer-facing copy for Fireball Softball Comforter Set; remove editorial/process wording, keep verified options/customizer labels, and preserve exact motif facts.
5. **MAJOR — product 334:** Cannibalization risk is not fully resolved even though motifs differ. Đề xuất: Sharpen secondary keyword map and internal linking around pink glove motif vs vintage/polka-dot/fireball variants.
6. **MAJOR — product 334:** This is not fully customer-facing publish-ready storefront copy even though the facts are largely verified. Đề xuất: Rewrite in English as customer-facing copy for Pink Glove Softball Comforter Set; remove editorial/process wording, keep verified options/customizer labels, and preserve exact motif facts.
7. **MAJOR — product 335:** This is not fully customer-facing publish-ready storefront copy even though the facts are largely verified. Đề xuất: Rewrite in English as customer-facing copy for Celtic Tree of Life Quilt Set; remove editorial/process wording, keep verified options/customizer labels, and preserve exact motif facts.
8. **MAJOR — product 336:** This is not fully customer-facing publish-ready storefront copy even though the facts are largely verified. Đề xuất: Rewrite in English as customer-facing copy for Winter Cardinal Berry Branch Quilt; remove editorial/process wording, keep verified options/customizer labels, and preserve exact motif facts.
9. **MAJOR — product 337:** The selected keyword may be too narrow or cannibalize adjacent products without enough demand proof. Đề xuất: Use a broader primary with the motif as secondary, or add stronger exact SERP/Search Console evidence.
10. **MAJOR — product 337:** This is not fully customer-facing publish-ready storefront copy even though the facts are largely verified. Đề xuất: Rewrite in English as customer-facing copy for Heart Branch Cardinals Christmas Quilt; remove editorial/process wording, keep verified options/customizer labels, and preserve exact motif facts.

## SERP và keyword

- Đã kiểm tra lại 16 truy vấn: mỗi sản phẩm có primary keyword và comparator gần nhất, giới hạn US/English, không claim volume/ranking.
- Nhóm softball 331–334 và cardinal 336–338 vẫn có nguy cơ overlap; các keyword exact rất niche được chấm PARTIAL khi SERP chỉ hỗ trợ broad/comparator intent.
- Chi tiết query/timestamp/URL đã đọc nằm trong `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260908_183309\serp_evidence.json`.

## Ảnh và alt text

- Đã mở trực tiếp đủ 57 ảnh trong run mới; không dùng contact sheet để thay thế việc xem ảnh.
- Alt/observation r3 khớp nội dung ảnh ở mức tốt. Product 334 có 14 ảnh, trong đó ảnh 8–14 là chuỗi lặp của 1–7; ghi MINOR để cân nhắc dọn gallery nếu không có nhu cầu app/variant.

## Đối chiếu r2 → r3

- 8 lỗi meta description bị cắt cụt trong `revision_summary.json` được đánh dấu **RESOLVED**: tất cả meta r3 kết thúc rõ nghĩa, không lơ lửng.
- Workbook nằm trong thư mục r3 và hash khớp, nhưng một số row-level metadata vẫn ghi `revision=r2`/contact sheet; ghi LIMITATION traceability, không coi là lỗi storefront.

## Giới hạn và trạng thái bàn giao

- Live HTML/product JSON đã được lưu lại; Python SSL verify báo lỗi nên HTML/JSON được tải bằng retry không xác minh SSL, có ghi limitation trong evidence.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại kiểm tra cấu trúc/công thức.
- Chỉ dừng ở batch 34. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA workbook: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\qa\20260908_183309\SEO_QA_qa_batch_034_r3.xlsx`
- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260908_183309\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260908_183309\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260908_183309\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260908_183309`
