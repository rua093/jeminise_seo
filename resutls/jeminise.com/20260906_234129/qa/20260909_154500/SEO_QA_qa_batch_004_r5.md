# SEO QA qa_batch_004_r5 - 20260909_154500

Chấm lại độc lập batch 04, phạm vi inventory 31-40, market US, language English. Workbook SEO nguồn và Shopify không bị sửa; `awaiting_confirmation=true`.

## Kết luận

Report mới không kế thừa điểm/status của run `20260909_043000`. Rating được ghi sau khi có evidence snapshot, live HTML/JSON, ảnh tải trực tiếp và kiểm tra SERP.

| Pos | Product ID | Final score | Status | Ghi chú chính |
|---:|---:|---:|---|---|
| 31 | 8834720170183 | 95.0 | QA_PASS | PASS đủ coverage |
| 32 | 8834709323975 | 95.0 | QA_PASS | PASS đủ coverage |
| 33 | 8867095806151 | 95.0 | QA_PASS | PASS đủ coverage |
| 34 | 8867095445703 | 77.5 | QA_REVISE | Cần sửa claim/intent |
| 35 | 8834768699591 | 85.0 | QA_REVISE | Cần runtime customizer / claim exact control |
| 36 | 8834754445511 | 85.0 | QA_REVISE | Cần runtime customizer / claim exact control |
| 37 | 8834752676039 | 85.0 | QA_REVISE | Cần runtime customizer / claim exact control |
| 38 | 8834767421639 | 85.0 | QA_REVISE | Cần runtime customizer / claim exact control |
| 39 | 8834773319879 | 85.0 | QA_REVISE | Cần runtime customizer / claim exact control |
| 40 | 8834768011463 | 85.0 | QA_REVISE | Cần runtime customizer / claim exact control |

## Các phát hiện chính

- 62/62 ảnh baseline được tải trực tiếp, mở ở độ phân giải đọc được và ghi `qa_image_key` duy nhất. Ảnh vẫn nằm trong mẫu số theo sản phẩm.
- 31/32 được tách theo birdhouse versus snowy branches; 33 xác nhận cow farmhouse landscape; 34 cần làm chặt crocodile thay vì dùng alligator/Dragonfly lẫn lộn.
- 35-40 không bị kết luận cannibalization chỉ vì cùng football/flag; modifier retro, grunge, cosmic, USA flag, paint splash và patriotic được đánh giá riêng.
- Với 35-40, claim personalization được chấm PARTIAL vì cần xác minh runtime customizer exact labels/limits; static HTML thiếu input không được dùng làm bằng chứng phủ định.
- E1 chỉ PARTIAL toàn lô vì revision metadata trong workbook/source summary còn dấu r4, nhưng không bị gán CRITICAL.

## Kiểm thử

- Counts: 10 product key, 110 criteria, 62 image rows, 40 Keyword_Map, 10 Buyer_Search_Research, 10 Product_Evidence.
- Logic tests: `100 + CRITICAL => QA_FAIL`; `90 + full coverage/no blocker => QA_PASS`; `72/80 => 72-92 + QA_INCOMPLETE`.
- Quét công thức không thấy `#REF!`, `#NAME?`, `#DIV/0!`; I1 lấy từ trung bình ảnh đúng product.

Output XLSX: `resutls\jeminise.com\20260906_234129\qa\20260909_154500\SEO_QA_qa_batch_004_r5.xlsx`
