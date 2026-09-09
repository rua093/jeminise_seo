# SEO Re-QA — qa_batch_001_r7

## Kết luận

- Phạm vi: **10 sản phẩm, 65/65 ảnh (100%)**; inventory position 1–10; revision **r7**.
- Điểm lô: **93.5/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 0 QA_FAIL, 3 QA_REVISE, 7 QA_PASS, 0 QA_INCOMPLETE.
- Phát hiện r7: 0 CRITICAL, 4 MAJOR, 2 MINOR, 12 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\revisions\qa_batch_001_r7\SEO_Product_Optimization_qa_batch_001_r7.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `A965CFD64371561152CF92B4B3B1CDF449E3F06BC3E0EB4060CAD54DC5AD713B`

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 1 | Cardinal Sunflower Autumn Quilt Set | 87.5 | QA_REVISE | 0/1/0/2 |
| 2 | Personalized Teal Softball Comforter Set | 97.5 | QA_PASS | 0/0/1/1 |
| 3 | Christmas Cardinal Memorial Quilt Set | 92.5 | QA_REVISE | 0/1/0/1 |
| 4 | Cardinal Roses Memorial Quilt Set | 85.0 | QA_REVISE | 0/2/0/1 |
| 5 | Colorful Cat Patchwork Quilt Set | 97.5 | QA_PASS | 0/0/0/1 |
| 6 | Geometric Cat Patchwork Quilt Set | 90.0 | QA_PASS | 0/0/0/2 |
| 7 | Celtic Fantasy Tree Quilt Set | 92.5 | QA_PASS | 0/0/0/1 |
| 8 | Celtic Tree of Life Quilt Set | 97.5 | QA_PASS | 0/0/0/1 |
| 9 | Farmhouse Chicken Patchwork Quilt Set | 97.5 | QA_PASS | 0/0/0/1 |
| 10 | Personalized God Says I Am Christian Bedding Set | 97.5 | QA_PASS | 0/0/1/1 |

## Phát hiện ưu tiên

1. **MAJOR — pos 1, 3, 4:** description r7 dùng kích thước từ panel ảnh cũ nhưng không khớp selector live: thiếu Full, sai Twin/Queen và đảo chiều King. Cần thay bằng đúng `Throw 60x70, Twin 68x86, Full 80x90, Queen 90x90, King 102x91`.
2. **MAJOR — pos 4:** câu `Use this product page for ... intent` là hướng dẫn SEO nội bộ, không phải copy cho khách hàng.
3. **Keyword:** pos 1 và 6 vẫn trung thực ở `HYPOTHESIS_ONLY`; pos 4 và một số exact long-tail chỉ được chấm PARTIAL vì SERP hỗ trợ broad/comparator intent tốt hơn exact phrase.
4. **MINOR — pos 2, 10:** H1/title nguồn live vẫn có ký tự thay thế `�`/từ bị cắt, dù title SEO r7 đã sạch.
5. **LIMITATION — cả 10:** file r7 và hash rõ ràng nhưng row-level `revision`/`revision_summary` vẫn ghi r5; đây là lỗi truy vết, không phải lỗi storefront.

## Ảnh, personalization và lịch sử r4

- Đã mở trực tiếp đủ 65 ảnh; 65 observation và alt `SET` của r7 đều khớp đúng cảnh/panel, I1 lấy từ điểm ảnh 100/100 của từng sản phẩm.
- Pos 2 có hai field bắt buộc `Custom Your Name` và `Custom Your Number`; pos 10 có field bắt buộc `Enter Name`, tối đa 13 ký tự. Claim r7 khớp control live.
- 23 issue r4 đã kiểm tra lại độc lập: 18 RESOLVED, 5 PERSISTS, 0 REGRESSED, 0 NOT_APPLICABLE. Chi tiết ở `history_issue_review.json`.
- Không trừ D1 chỉ vì meta description nằm ngoài 145–165 ký tự; tất cả meta r7 đều hoàn chỉnh và rõ nghĩa.

## Giới hạn và bàn giao

- Live identity, canonical, H1, product ID và gallery count khớp workbook; khác biệt entity HTML ở meta pos 5–6 không phải drift ngữ nghĩa.
- Không sửa workbook r7, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Đúng 5 sheet, filter/freeze/wrap/hyperlink và công thức truy kiểm; preview được tạo ngoài project.
- Dừng sau batch 01. `awaiting_confirmation=true`.
