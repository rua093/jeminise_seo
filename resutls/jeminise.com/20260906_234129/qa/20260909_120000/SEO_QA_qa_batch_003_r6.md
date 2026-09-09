# QA độc lập `qa_batch_003_r6`

- `qa_run_id`: `20260909_120000`
- Market/language: US / English
- Source SHA-256: `95AF231168F94DD35DBCB3B374F40DB2CE7EE17A768C22FF0FB1D99BE48C5C5A` (trước và sau QA không đổi)
- Scope: inventory positions 21-30 khóa bằng `product_key`; 10 products / 50 images.
- Live evidence: 10 HTML + 10 product JSON reads; 50 direct image downloads at 1200x1200; 20 US-English SERP queries from the prior evidence set are re-linked for review.

## Kết quả

| Pos | Product ID | Score | Status | Finding |
|---:|---:|---:|---|---|
| 21 | 8867157835975 | 90.0 | QA_PASS | Live personalization control reconciles |
| 22 | 8867157868743 | 67.5 | QA_FAIL | Unsupported/contradictory personalization claim |
| 23 | 8867157999815 | 67.5 | QA_FAIL | Unsupported/contradictory personalization claim |
| 24 | 8867158393031 | 67.5 | QA_FAIL | Unsupported/contradictory personalization claim |
| 25 | 8867157737671 | 67.5 | QA_FAIL | Unsupported/contradictory personalization claim |
| 26 | 8867158261959 | 67.5 | QA_FAIL | Unsupported/contradictory personalization claim |
| 27 | 8867157803207 | 67.5 | QA_FAIL | Unsupported/contradictory personalization claim |
| 28 | 8867157704903 | 67.5 | QA_FAIL | Unsupported/contradictory personalization claim |
| 29 | 8834708111559 | 85.0 | QA_REVISE | Live control exists but r6 says no field |
| 30 | 8834727805127 | 85.0 | QA_REVISE | Live control exists but r6 says no field |

## Nhận xét chính

- Ảnh 21 và 24 được phân biệt đúng: candy cane gingerbread versus gingerbread village.
- Ảnh 22 và 28 được phân biệt đúng: black Christmas tree versus vintage Christmas tree.
- Ảnh 23/25/27 xác nhận birdhouse/cardinal, một snowman và hai snowmen; không thấy trùng motif trong gallery.
- Ảnh 29 có chữ “I am Always With You” và branch/cardinal memorial; ảnh 30 có cardinal với wreath/pine/holly/pinecones và sample “Sophia”.
- `P1/P2/T1/D2`: r6 claim “Personalized”/name field không có live control ở 22-28; r6 lại phủ nhận control ở 29-30 dù live JSON/body có `1 text input`. Đây là MAJOR.
- D2 đã loại câu quy trình/QA/copy-about-copy cũ trong r6; lịch sử được ghi `RESOLVED`, không lấy điểm r4.

## Kiểm thử logic

- 100 điểm + CRITICAL → `QA_FAIL`.
- 90 điểm, đủ coverage, không blocker → `QA_PASS`.
- 72/80 assessed weight → khoảng điểm 72-92 và `QA_INCOMPLETE`.
- Workbook có đúng 5 sheet; 10 product rows; 110 criteria rows; 50 image rows; issue IDs và `qa_image_key` sinh duy nhất; formula ranges trỏ tới QA sheets.

## Bàn giao

- Workbook: `SEO_QA_qa_batch_003_r6.xlsx`
- `awaiting_confirmation=true`; giữ QA r4 cũ vì QA mới có finding cần sửa; không tạo APPROVED/import và không chuyển batch 04.
