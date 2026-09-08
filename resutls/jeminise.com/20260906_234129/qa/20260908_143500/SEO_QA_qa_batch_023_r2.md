# Báo cáo QA Độc lập — Batch 23 Revision r2 (`qa_batch_023_r2`)

**Mã phiên QA (qa_run_id):** `20260908_143500`  
**Thời gian thực hiện:** 2026-09-08 14:35:00 (GMT+7)  
**Người thực hiện (Auditor):** Antigravity Independent QA Assistant  
**Thị trường mục tiêu (Market):** United States (US) | **Ngôn ngữ SEO:** English  
**Trạng thái kiểm định tổng thể:** `QA_REVISE` (Cần chỉnh sửa trước khi duyệt xuất bản)  

---

## 1. Nguồn dữ liệu, Phạm vi & Tính Toàn vẹn Dữ liệu

- **Workbook nguồn r2:** `resutls/jeminise.com/20260906_234129/revisions/qa_batch_023_r2/SEO_Product_Optimization_qa_batch_023_r2.xlsx`
- **SHA-256 Workbook nguồn:** `7E965F336D336B5D6AEBE2CBCC508865569E32315DB6619410E2DF9A22BED809` *(Khớp tuyệt đối 100% với hash đóng băng trong yêu cầu)*
- **Snapshot lưu trữ:** `seo_runs/jeminise.com/20260906_234129/qa/20260908_143500/source_snapshot/SEO_Product_Optimization_qa_batch_023_r2.xlsx`
- **SHA-256 Snapshot:** `7E965F336D336B5D6AEBE2CBCC508865569E32315DB6619410E2DF9A22BED809` *(Đã kiểm tra khớp bit-for-bit)*
- **Phạm vi kiểm tra:** Cố định đúng 10 sản phẩm (Inventory Positions 221 – 230), 70 hình ảnh gallery, 40 dòng Keyword_Map và 10 dòng Buyer_Search_Research.
- **Tình trạng QA trước đây:** Xác nhận **đây là lượt QA đầu tiên cho batch 23 r2**. Không tìm thấy run QA cũ nào; không tạo giả định lịch sử issue.
- **Đối chiếu Admin Export Baseline:**
  - Bản r2 viện dẫn hash lịch sử: `729CBD9B3E0248F5D9196D999A49222DBDACA85C2F3CF8524BEEC549B57D738C` (`729CBD9`).
  - File `products_export_1.csv` hiện hành trong workspace có hash: `97AA8DC283CFC927BF3B6F41A3EA8926C96C0723B6942CF7C1D147CD854FBA83` (`97AA8DC`).
  - Đã rà soát lịch sử Git (`git log`); commit duy nhất chứa `products_export_1.csv` là `f3e38d5` mang hash `97AA8DC`. QA ghi nhận đây là **giới hạn đối chiếu lịch sử** (documented limitation), không mặc định bản r2 sai, và đã trực tiếp đối chiếu live storefront JSON, page HTML snapshot và 70 file ảnh local thay thế.

### Bảng Tổng hợp Chỉ số KPI Kiểm toán Lô

| Chỉ số kiểm toán (Metric) | Giá trị ghi nhận | Quy chuẩn Rubric | Đánh giá |
| :--- | :---: | :---: | :---: |
| Tổng số sản phẩm kiểm định | **10 / 10** | 10 sản phẩm/lô | Đạt 100% phạm vi |
| Tỷ lệ đọc trang (Page Read) | **100% (10/10)** | Bắt buộc 100% | Đạt |
| Tổng số ảnh kiểm toán trực tiếp | **70 / 70** | 100% ảnh gallery | Đạt 100% phạm vi |
| Trọng số đánh giá (Assessed Weight) | **100.0 / 100** | 100.0 | Đầy đủ dữ liệu |
| Điểm số trung bình toàn lô | **95.0 / 100** | Thang điểm 100 | Xuất sắc về kỹ thuật |
| Trạng thái từng sản phẩm | **10 QA_REVISE** | Ngưỡng 85đ & không có MAJOR | Cần sửa 2 lỗi nội dung |
| Số lỗi nghiêm trọng (CRITICAL) | **0** | 0 | Không có lỗi chặn |
| Số lỗi chính (MAJOR) | **10** (1 lỗi/sản phẩm tại D2) | 0 để đạt QA_PASS | Cần khắc phục |
| Số lỗi nhỏ (MINOR) | **10** (1 lỗi/sản phẩm tại D1) | Khắc phục để hoàn thiện | Cần khắc phục |
| Giới hạn nguồn (LIMITATION) | **10** (Đối chiếu export lịch sử) | Ghi nhận minh bạch | Đã lập hồ sơ |

---

## 2. Bảng Điểm Chi tiết Từng Sản phẩm (10 Sản phẩm)

| Pos | Handle | Title đề xuất r2 | Primary Keyword | P1 | P2 | K1 | K2 | K3 | T1 | T2 | D1 | D2 | I1 | E1 | Final Score | Lỗi ghi nhận | Kết luận QA |
|:---:|:---|:---|:---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---|:---:|
| 221 | `personalized-basketball-hoop-with-net-close-up-blanket-with-name-and-number-cc637164fc-cc637164fc` | Custom Basketball Hoop Blanket | *custom basketball hoop blanket* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 222 | `personalized-basketball-next-to-athletic-shoes-blanket-with-name-and-number-dfbf4fde2f-dfbf4fde2f` | Custom Basketball Shoes Blanket | *custom basketball shoes blanket* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 223 | `personalized-basketball-on-court-lines-comforter-name-number-303fbd50f6-303fbd50f6` | Custom Court Lines Basketball Comforter | *custom court lines basketball comforter* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 224 | `personalized-basketball-on-court-perspective-comforter-with-name-and-number-1acfbe4260-1acfbe4260` | Custom Basketball Court Perspective Comforter | *custom basketball court perspective comforter* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 225 | `personalized-basketball-on-hardwood-court-comforter-with-name-and-number-8aee799947-8aee799947` | Custom Hardwood Court Basketball Comforter | *custom hardwood court basketball comforter* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 226 | `personalized-neon-basketball-player-blanket-788fb6ae63-788fb6ae63` | Custom Neon Basketball Player Blanket | *custom neon basketball player blanket* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 227 | `personalized-basketball-player-front-blanket-with-name-and-number-258e7f9e2e-258e7f9e2e` | Custom Flaming Basketball Player Blanket | *custom flaming basketball player blanket* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 228 | `personalized-basketball-player-silhouette-blanket-with-name-and-number-35f8822fba-35f8822fba` | Custom Dribbling Silhouette Basketball Blanket | *custom dribbling silhouette basketball blanket* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 229 | `personalized-basketball-players-and-hoops-comforter-with-name-and-number-ff13bad888-ff13bad888` | Custom Basketball Collage Comforter | *custom basketball collage comforter* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 230 | `personalized-basketball-shattered-glass-comforter-with-name-and-number-353fc0c17e-353fc0c17e` | Custom Shattered Glass Basketball Comforter | *custom shattered glass basketball comforter* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |

---

## 3. Phân tích Chi tiết 11 Tiêu chuẩn Đánh giá & Các điểm Đặc thù Batch 23

### 3.1. Nhóm Tiêu chuẩn Đạt Điểm Tuyệt đối (FULL)
- **P1. Đúng sản phẩm và thiết kế (15/15đ):** Phân biệt rành mạch hai nhóm form factor thể thao: 5 sản phẩm Chăn (Blanket - Pos 221, 222, 226, 227, 228) và 5 sản phẩm Chăn ga/Bộ chăn (Comforter / Duvet Cover - Pos 223, 224, 225, 229, 230). Nhận diện chính xác 100% các motif đồ họa độc lập:
  - *Pos 221:* Cận cảnh rổ bóng rổ phát sáng và lưới trắng nổi bật (`hoop and net close-up`).
  - *Pos 222:* Quả bóng rổ đặt cạnh giày thể thao thi đấu đã sờn (`basketball next to athletic shoes`).
  - *Pos 223:* Vạch sân bóng rổ vàng kim trên nền kết cấu sân nứt 3D (`gold court lines & cracked texture`).
  - *Pos 224:* Phối cảnh góc nghiêng sân đấu sắc thái đen-đỏ với bóng rổ lớn (`black-red court perspective`).
  - *Pos 225:* Sân đấu bóng rổ sàn gỗ arena phối màu xanh-vàng kim (`blue-gold hardwood court`). *Lưu ý: 'hardwood' là đồ họa in 3D, không phải chất liệu gỗ vật lý của sản phẩm.*
  - *Pos 226:* Hình bóng cầu thủ úp rổ phong cách neon phát sáng kèm rổ (`neon dunking player silhouette`).
  - *Pos 227:* Cầu thủ dẫn bóng chính diện bốc lửa với hiệu ứng ngọn lửa rực cháy (`front-facing dribbler & flame effect`).
  - *Pos 228:* Bóng đen cầu thủ dẫn bóng phong cách tối giản trên nền xám-cam với quả bóng phóng to (`black dribbling silhouette & oversized ball`).
  - *Pos 229:* Tranh ghép collage đa họa tiết gồm bóng, rổ và bóng vận động viên (`basketball collage comforter`).
  - *Pos 230:* Hiệu ứng kính vỡ 3D văng mảnh quanh quả bóng rổ trung tâm (`shattered glass basketball comforter`).
- **P2. Đúng thuộc tính và tùy chọn cá nhân hóa (10/10đ):** Kiểm toán cấu hình thực tế trên live storefront và mã ứng dụng customizer:
  - *Nhóm Blanket (Pos 221, 222, 226, 227, 228):* Có **2 trường TÙY CHỌN** (`required: false`): `Custom Name` (1–200 ký tự, placeholder `David`) và `Custom Number` (1–20 ký tự, placeholder `22`). Nhóm tùy chọn duy nhất là `Choose Your Size` (8 biến thể gồm 4 kích cỡ Fleece và 4 kích cỡ Sherpa).
  - *Nhóm Comforter (Pos 223, 224, 225, 229, 230):* Có **2 trường BẮT BUỘC** (`required: true`): `Customize Your Name` (1–30 ký tự) và `Customize Your Number` (1–5 ký tự, placeholder `10`). Có 3 nhóm tùy chọn gồm `Choose Product Type + Size` (8 lựa chọn Duvet Cover / Comforter x 4 kích thước Twin/Full/Queen/King), `Choose Pillowcases` (3 lựa chọn: None / 1 / 2 vỏ gối), và `Additional Sheet Cover (Flat Sheet)` (2 lựa chọn: None / 1 ga trải giường cùng size), tạo nên tổng cộng 48 biến thể.
- **K1, K2, K3. Bộ Từ khóa & Nhu cầu Tìm kiếm (20/20đ):** Bộ từ khóa phân bổ chuẩn xác mục đích mua sắm tại Mỹ, phân định rạch ròi giữa chăn đắp (`blanket/throw`) và bộ chăn ga giường (`comforter/duvet cover`). Mức chứng thực ghi nhận trung thực là `SERP_ONLY` vì nguồn dựa trên kết quả SERP thực tế của Google, Amazon, Etsy và Zazzle mà không giả định số liệu volume trả phí.
- **T1, T2. Tiêu đề SEO & Product Title (15/15đ):** Tiêu đề ngắn gọn, tự nhiên, chứa từ khóa chính ngay đầu; Meta Title đạt chuẩn dưới 60 ký tự.
- **I1. Kiểm toán Trực quan Toàn bộ 70 Ảnh (20/20đ):** Tất cả 70 ảnh đã được xem trực tiếp ở độ phân giải pixel gốc; điểm trung bình đạt 100/100đ, đóng góp trọn vẹn 20.0 điểm vào điểm sản phẩm.
- **E1. Tính Toàn vẹn Thực chứng (5/5đ):** Đầy đủ hồ sơ thực chứng: file ảnh local, cấu hình customizer, đối chiếu live storefront và 20 truy vấn SERP độc lập.

### 3.2. Tiêu chuẩn Bị Trừ điểm: D1. Meta Description SEO (3/5đ — Đạt mức PARTIAL)
- **Hiện trạng kiểm toán:** Trong bản r2, **toàn bộ 10 Meta Description bị cắt cụt cơ học ở đúng 155 ký tự** (`hard-truncated at 155 chars`), dẫn tới việc đứt ngang câu hoặc cắt cụt từ ngữ:
  - **Pos 221:** kết thúc bằng `...selectable` (thiếu danh từ và dấu chấm kết câu).
  - **Pos 222:** kết thúc bằng `...selectab` (cắt cụt từ `selectable`).
  - **Pos 223:** kết thúc bằng `...sample name and jersey numbe` (cắt cụt từ `number`).
  - **Pos 224:** kết thúc bằng `...name and jersey number, s` (cắt cụt từ `selectable`).
  - **Pos 225:** kết thúc bằng `...nd jersey number, selecta` (cắt cụt từ `selectable`).
  - **Pos 226:** kết thúc bằng `...ey number, selectable siz` (cắt cụt từ `size` hoặc `sizes`).
  - **Pos 227:** kết thúc bằng `...name and jersey number, ` (dấu phẩy và khoảng trắng lơ lửng cuối câu).
  - **Pos 228:** kết thúc bằng `...oversized ball, sample name and ` (từ nối `and` lơ lửng cuối câu).
  - **Pos 229:** kết thúc bằng `...sample name and jersey ` (cắt lửng cụm từ `jersey number`).
  - **Pos 230:** kết thúc bằng `...nd jersey number on pillo` (cắt cụt từ `pillow` / `pillow shams`).
- **Xử lý:** Trừ 2 điểm D1 (đạt 3/5đ, PARTIAL), ghi nhận 10 lỗi MINOR (`ISSUE-0011` đến `ISSUE-0020`), và cung cấp 10 câu Meta Description hoàn chỉnh chuẩn 151–160 ký tự tại Mục 6.1.

### 3.3. Tiêu chuẩn Bị Trừ điểm: D2. Product Description HTML (7/10đ — Đạt mức PARTIAL)
- **Hiện trạng kiểm toán:**
  1. **Lộ câu lệnh sinh nội bộ (Prompt boilerplate leakage):** Cả 10 mô tả r2 chứa các câu văn sinh nội bộ hiển thị trực tiếp cho người mua hàng:
     > *"The copy stays specific to the visible basketball artwork, product form and selectable options for this exact item."*  
     > *"Gallery images include the main mockup plus size, feature, care, bedding-type or lifestyle panels where shown."*
  2. **Hướng dẫn chọn biến thể mâu thuẫn đối với nhóm Blanket (Pos 221, 222, 226, 227, 228):**
     - Bản r2 viết: *"Select the product type and size shown on the product page before checkout."*
     - Thực tế sản phẩm Blanket chỉ có duy nhất 1 nhóm tùy chọn là `Choose Your Size` (không có tùy chọn `Product Type`). Hướng dẫn này bị sao chép nhầm từ nhóm Comforter.
  3. **Đặc thù Sản phẩm Pos 230 (`353fc0c17e` - Shattered Glass Comforter):**
     - Bản r2 đề cập tên/số trên `pillow shams` mà không làm rõ điều kiện mua hàng. Trong thực tế, nhóm tùy chọn `Choose Pillowcases` là tùy chọn mua thêm (`None`, `1 Pillowcase`, `2 Pillowcases`). Nếu khách chọn `None`, họ sẽ không nhận được vỏ gối.
  4. **Thiếu sót thông số kỹ thuật thiết yếu:** Thiếu phân định rõ ràng giữa cấu tạo Duvet Cover (vỏ chăn có khóa kéo ẩn và dây buộc góc) vs Comforter (chăn chần bông đệm trọn bộ); thiếu hướng dẫn giặt ủi và bảng kích thước chi tiết.
- **Xử lý:** Trừ 3 điểm D2 (đạt 7/10đ, PARTIAL), ghi nhận 10 lỗi MAJOR (`ISSUE-0001` đến `ISSUE-0010`), và cung cấp 10 bản mô tả HTML publish-ready thay thế tại Mục 6.2.

---

## 4. Kết quả Kiểm toán Trực quan Toàn bộ 70 Hình ảnh Gallery

Tất cả 70 hình ảnh đã được mở trực tiếp ở độ phân giải gốc để kiểm tra chi tiết đồ họa:

| Pos | Ảnh # | Tên file local | Media ID | Kích thước | SHA-256 (rút gọn) | Chi tiết kiểm toán trực quan (Direct Visual Audit) | Alt Text r2 | IM1 | IM2 | IM3 | IM4 | Điểm |
|:---:|:---:|:---|:---:|:---:|:---:|:---|:---|:---:|:---:|:---:|:---:|:---:|
| 221 | 1 | `221_01.jpg` | 47538875728071 | 1000x1000 | `3dd9f79469...` | Blanket flat mockup with close-up hoop, white net, sample name Justin and number 23. | Custom basketball hoop blanket flat view | 40 | 30 | 20 | 10 | **100** |
| 221 | 2 | `221_02.jpg` | 47538875760839 | 1000x1000 | `36fbbe5dc3...` | Person holds the basketball hoop blanket in a living room with the same sample name and number. | Basketball hoop blanket held in living room | 40 | 30 | 20 | 10 | **100** |
| 221 | 3 | `221_03.jpg` | 47538875793607 | 1000x1000 | `41a49d1b14...` | Blanket size chart showing 40x30 through 80x60 options with scale figure. | Basketball hoop blanket size chart | 40 | 30 | 20 | 10 | **100** |
| 221 | 4 | `221_04.jpg` | 47538875826375 | 1000x1000 | `8331d743bc...` | Basketball hoop blanket draped across a sofa under a window. | Basketball hoop blanket on sofa | 40 | 30 | 20 | 10 | **100** |
| 221 | 5 | `221_05.jpg` | 47538875859143 | 1000x1000 | `6a4988d239...` | Feature panel overlays close-up blanket artwork with fluffy, quality and no-shedding callouts. | Basketball hoop blanket feature panel | 40 | 30 | 20 | 10 | **100** |
| 221 | 6 | `221_06.jpg` | 47538875891911 | 1000x1000 | `72b454fb10...` | Folded blanket and fabric/care collage with breathable, skin-friendly and machine washable callouts. | Basketball hoop blanket fabric care panel | 40 | 30 | 20 | 10 | **100** |
| 221 | 7 | `221_07.jpg` | 47538875924679 | 1000x1000 | `b74f4112ed...` | Custom blanket size guide showing bed placement examples and width/length table. | Basketball hoop blanket bed size guide | 40 | 30 | 20 | 10 | **100** |
| 221 | 8 | `221_08.jpg` | 47538875957447 | 1000x1000 | `7fd39fe1be...` | Lifestyle image with adult and child reading under the basketball hoop blanket. | Basketball hoop blanket family lifestyle image | 40 | 30 | 20 | 10 | **100** |
| 222 | 1 | `222_01.jpg` | 47538986451143 | 1000x1000 | `44057120f5...` | Blanket flat mockup with dark basketball and athletic shoes artwork, sample name Emery and number 33. | Custom basketball shoes blanket flat view | 40 | 30 | 20 | 10 | **100** |
| 222 | 2 | `222_02.jpg` | 47538986483911 | 1000x1000 | `e856eeb24a...` | Person holds basketball shoes blanket in a living room setting. | Basketball shoes blanket held in living room | 40 | 30 | 20 | 10 | **100** |
| 222 | 3 | `222_03.jpg` | 47538986516679 | 1000x1000 | `8abc316aab...` | Size chart panel showing four blanket sizes with basketball shoes artwork. | Basketball shoes blanket size chart | 40 | 30 | 20 | 10 | **100** |
| 222 | 4 | `222_04.jpg` | 47538986549447 | 1000x1000 | `6a97e4d058...` | Feature panel over ball-and-shoes artwork with fluffy, high-quality and no-pilling callouts. | Basketball shoes blanket feature panel | 40 | 30 | 20 | 10 | **100** |
| 222 | 5 | `222_05.jpg` | 47538986582215 | 1000x1000 | `384ccf52a1...` | Basketball shoes blanket draped on sofa beneath round wall mirror. | Basketball shoes blanket on sofa | 40 | 30 | 20 | 10 | **100** |
| 222 | 6 | `222_06.jpg` | 47538986614983 | 1000x1000 | `abf5626489...` | Folded blanket and fabric/care collage for the basketball shoes design. | Basketball shoes blanket fabric care panel | 40 | 30 | 20 | 10 | **100** |
| 222 | 7 | `222_07.jpg` | 47538986647751 | 1000x1000 | `8931925a2f...` | Custom blanket bed size guide with width and length table. | Basketball shoes blanket bed size guide | 40 | 30 | 20 | 10 | **100** |
| 222 | 8 | `222_08.jpg` | 47538986680519 | 1000x1000 | `3d1b6e821b...` | Adult and child reading under a blanket showing the script sample name. | Basketball shoes blanket family lifestyle image | 40 | 30 | 20 | 10 | **100** |
| 223 | 1 | `223_01.jpg` | 47538295668935 | 1000x1000 | `02ee999bae...` | Bedroom mockup of black court-lines basketball comforter with sample name T Mack and number 5. | Custom court lines basketball comforter | 40 | 30 | 20 | 10 | **100** |
| 223 | 2 | `223_02.jpg` | 47538295701703 | 1000x1000 | `701693c984...` | Second bedroom mockup of the court-lines comforter with folded white duvet edge. | Court lines basketball bedding mockup | 40 | 30 | 20 | 10 | **100** |
| 223 | 3 | `223_03.jpg` | 47538295734471 | 1000x1000 | `8b3c12074c...` | Panel comparing duvet cover set and comforter set for all-season use. | Basketball bedding type comparison panel | 40 | 30 | 20 | 10 | **100** |
| 223 | 4 | `223_04.jpg` | 47538295767239 | 1000x1000 | `8ec870ac19...` | Feature panel with basketball court print close-ups, microfiber and 3D printed pattern callouts. | Court lines basketball comforter feature panel | 40 | 30 | 20 | 10 | **100** |
| 223 | 5 | `223_05.jpg` | 47538295800007 | 1000x1000 | `5d472e3947...` | Low-angle bed mockup with court-lines basketball bedding and comfort icon strip. | Court lines basketball bedding close view | 40 | 30 | 20 | 10 | **100** |
| 223 | 6 | `223_06.jpg` | 47538295832775 | 1000x1000 | `e44cab8b26...` | Easy care panel with pillows and wrinkle-free, stain-proof, anti-pilling callouts. | Basketball bedding easy care panel | 40 | 30 | 20 | 10 | **100** |
| 223 | 7 | `223_07.jpg` | 47538295865543 | 1000x1000 | `2f82ab909c...` | Size dimension panel showing twin, full, queen and king bed sizes. | Basketball comforter size dimension panel | 40 | 30 | 20 | 10 | **100** |
| 224 | 1 | `224_01.jpg` | 47538203951303 | 1000x1000 | `9f76e36c0f...` | Bedroom mockup of black and red basketball court perspective comforter with sample name Anthony and number 8. | Custom basketball court perspective comforter | 40 | 30 | 20 | 10 | **100** |
| 224 | 2 | `224_02.jpg` | 47538203984071 | 1000x1000 | `dd43faa377...` | White zipper close-up panel labeled bottom zippered closure. | Bottom zippered closure bedding panel | 40 | 30 | 20 | 10 | **100** |
| 224 | 3 | `224_03.jpg` | 47538204016839 | 1000x1000 | `0b6cd797d6...` | High-density weaving panel with fabric icons and weave comparison. | High-density weaving bedding panel | 40 | 30 | 20 | 10 | **100** |
| 224 | 4 | `224_04.jpg` | 47538204049607 | 1000x1000 | `9da7310b60...` | Machine washable panel with washing machine and laundry baskets. | Machine washable bedding care panel | 40 | 30 | 20 | 10 | **100** |
| 224 | 5 | `224_05.jpg` | 47538204082375 | 1000x1000 | `a88ad14e34...` | Blue table size panel for duvet cover and pillowcase dimensions with bed icons. | Basketball bedding size chart panel | 40 | 30 | 20 | 10 | **100** |
| 225 | 1 | `225_01.jpg` | 47538241732807 | 1000x1000 | `8fe1d17990...` | Bedroom mockup of blue and gold hardwood court basketball comforter with sample name Matthew and number 2. | Custom hardwood court basketball comforter | 40 | 30 | 20 | 10 | **100** |
| 225 | 2 | `225_02.jpg` | 47538241765575 | 1000x1000 | `dd43faa377...` | White zipper close-up panel labeled bottom zippered closure. | Bottom zippered closure bedding panel for hardwood court comforter | 40 | 30 | 20 | 10 | **100** |
| 225 | 3 | `225_03.jpg` | 47538241798343 | 1000x1000 | `0b6cd797d6...` | High-density weaving feature panel with fabric and breathability icons. | High-density weaving panel for hardwood basketball bedding | 40 | 30 | 20 | 10 | **100** |
| 225 | 4 | `225_04.jpg` | 47538241831111 | 1000x1000 | `9da7310b60...` | Machine washable care panel with washing machine and laundry baskets. | Hardwood basketball bedding easy care panel | 40 | 30 | 20 | 10 | **100** |
| 225 | 5 | `225_05.jpg` | 47538241863879 | 1000x1000 | `a88ad14e34...` | Size chart panel listing US twin, full, queen and king duvet cover and pillowcase dimensions. | Hardwood basketball bedding size chart | 40 | 30 | 20 | 10 | **100** |
| 226 | 1 | `226_01.jpg` | 47539119587527 | 1000x1000 | `babdd95b0c...` | Blanket flat mockup with neon dunking player silhouette, hoop, sample name Fernando and number 51. | Custom neon basketball player blanket flat view | 40 | 30 | 20 | 10 | **100** |
| 226 | 2 | `226_02.jpg` | 47539119620295 | 1000x1000 | `51a37e938a...` | Person holds neon basketball player blanket in a living room. | Neon basketball player blanket held in living room | 40 | 30 | 20 | 10 | **100** |
| 226 | 3 | `226_03.jpg` | 47539119653063 | 1000x1000 | `99e423457b...` | Size chart panel showing four blanket sizes for the dunking player design. | Neon basketball player blanket size chart | 40 | 30 | 20 | 10 | **100** |
| 226 | 4 | `226_04.jpg` | 47539119685831 | 1000x1000 | `30532ad84f...` | Feature panel over player-and-hoop artwork with fluffy and no-pilling callouts. | Neon basketball player blanket feature panel | 40 | 30 | 20 | 10 | **100** |
| 226 | 5 | `226_05.jpg` | 47539119718599 | 1000x1000 | `87d0445cc8...` | Neon basketball player blanket draped across a sofa. | Neon basketball player blanket on sofa | 40 | 30 | 20 | 10 | **100** |
| 226 | 6 | `226_06.jpg` | 47539119751367 | 1000x1000 | `ef7669eca1...` | Folded blanket and fabric/care collage for the neon player design. | Neon basketball player blanket fabric care panel | 40 | 30 | 20 | 10 | **100** |
| 226 | 7 | `226_07.jpg` | 47539119784135 | 1000x1000 | `5cbcc0d92b...` | Custom blanket bed size guide with example placements and size table. | Neon basketball player blanket bed size guide | 40 | 30 | 20 | 10 | **100** |
| 226 | 8 | `226_08.jpg` | 47539119816903 | 1000x1000 | `38105ba4a7...` | Adult and child reading under the neon basketball player blanket. | Neon basketball player blanket family lifestyle image | 40 | 30 | 20 | 10 | **100** |
| 227 | 1 | `227_01.jpg` | 47539098747079 | 1000x1000 | `93768dfd49...` | Blanket flat mockup with front dribbling basketball player, flame effect, sample name King David and number 23. | Custom flaming basketball player blanket flat view | 40 | 30 | 20 | 10 | **100** |
| 227 | 2 | `227_02.jpg` | 47539098779847 | 1000x1000 | `daed17e5ab...` | Person holds flaming basketball player blanket in a living room. | Flaming basketball player blanket held in living room | 40 | 30 | 20 | 10 | **100** |
| 227 | 3 | `227_03.jpg` | 47539098812615 | 1000x1000 | `f71cb1660b...` | Blanket size chart with four sizes and player artwork thumbnails. | Flaming basketball player blanket size chart | 40 | 30 | 20 | 10 | **100** |
| 227 | 4 | `227_04.jpg` | 47539098845383 | 1000x1000 | `4dd212058a...` | Feature panel over flame player artwork with fluffy, quality and no-shedding callouts. | Flaming basketball player blanket feature panel | 40 | 30 | 20 | 10 | **100** |
| 227 | 5 | `227_05.jpg` | 47539098878151 | 1000x1000 | `bfef387b5f...` | Flaming basketball player blanket draped over sofa. | Flaming basketball player blanket on sofa | 40 | 30 | 20 | 10 | **100** |
| 227 | 6 | `227_06.jpg` | 47539098910919 | 1000x1000 | `e469c9731d...` | Folded blanket and fabric/care collage for the flaming player design. | Flaming basketball player blanket fabric care panel | 40 | 30 | 20 | 10 | **100** |
| 227 | 7 | `227_07.jpg` | 47539098943687 | 1000x1000 | `3f97f92e20...` | Custom blanket bed size guide with width and length table. | Flaming basketball player blanket bed size guide | 40 | 30 | 20 | 10 | **100** |
| 227 | 8 | `227_08.jpg` | 47539098976455 | 1000x1000 | `ba630f52ce...` | Adult and child reading under blanket showing flame player artwork. | Flaming basketball player blanket family lifestyle image | 40 | 30 | 20 | 10 | **100** |
| 228 | 1 | `228_01.jpg` | 47538931728583 | 1000x1000 | `762c5b9de3...` | Blanket flat mockup with black dribbling player silhouette, orange ball graphic, sample name Connor and number 3. | Custom dribbling silhouette basketball blanket flat view | 40 | 30 | 20 | 10 | **100** |
| 228 | 2 | `228_02.jpg` | 47538931761351 | 1000x1000 | `33677cdf81...` | Person holds dribbling silhouette basketball blanket in a living room. | Dribbling silhouette basketball blanket held in living room | 40 | 30 | 20 | 10 | **100** |
| 228 | 3 | `228_03.jpg` | 47538931794119 | 1000x1000 | `383a5dac27...` | Size chart panel showing four blanket sizes with silhouette artwork thumbnails. | Dribbling silhouette basketball blanket size chart | 40 | 30 | 20 | 10 | **100** |
| 228 | 4 | `228_04.jpg` | 47538931826887 | 1000x1000 | `e8db183eef...` | Feature panel over gray and orange basketball artwork with fabric callouts. | Dribbling silhouette basketball blanket feature panel | 40 | 30 | 20 | 10 | **100** |
| 228 | 5 | `228_05.jpg` | 47538931859655 | 1000x1000 | `817f4deb94...` | Dribbling silhouette basketball blanket draped on sofa. | Dribbling silhouette basketball blanket on sofa | 40 | 30 | 20 | 10 | **100** |
| 228 | 6 | `228_06.jpg` | 47538931892423 | 1000x1000 | `7cdb1f6266...` | Folded blanket and fabric/care collage for the silhouette basketball design. | Dribbling silhouette basketball blanket fabric care panel | 40 | 30 | 20 | 10 | **100** |
| 228 | 7 | `228_07.jpg` | 47538931925191 | 1000x1000 | `c9274ce258...` | Custom blanket bed size guide with width and length table. | Dribbling silhouette basketball blanket bed size guide | 40 | 30 | 20 | 10 | **100** |
| 228 | 8 | `228_08.jpg` | 47538931957959 | 1000x1000 | `ae9f638d10...` | Adult and child reading under blanket showing dribbling silhouette artwork. | Dribbling silhouette basketball blanket family lifestyle image | 40 | 30 | 20 | 10 | **100** |
| 229 | 1 | `229_01.jpg` | 47538193858759 | 1000x1000 | `361a9bc717...` | Bedroom mockup of basketball collage comforter with balls, hoops, player silhouettes, sample name Bray and number 12. | Custom basketball collage comforter | 40 | 30 | 20 | 10 | **100** |
| 229 | 2 | `229_02.jpg` | 47538193891527 | 1000x1000 | `c1bd1e723f...` | Second bed mockup of basketball collage bedding with folded white duvet edge. | Basketball collage bedding mockup | 40 | 30 | 20 | 10 | **100** |
| 229 | 3 | `229_03.jpg` | 47538193924295 | 1000x1000 | `0da7b7c24a...` | Duvet cover set versus comforter set comparison panel. | Basketball collage bedding type panel | 40 | 30 | 20 | 10 | **100** |
| 229 | 4 | `229_04.jpg` | 47538193957063 | 1000x1000 | `217e97712c...` | Feature panel with collage print close-ups, microfiber and 3D printed pattern callouts. | Basketball collage comforter feature panel | 40 | 30 | 20 | 10 | **100** |
| 229 | 5 | `229_05.jpg` | 47538193989831 | 1000x1000 | `637da24b84...` | Low-angle bed mockup with basketball collage print and comfort icon strip. | Basketball collage bedding close view | 40 | 30 | 20 | 10 | **100** |
| 229 | 6 | `229_06.jpg` | 47538194022599 | 1000x1000 | `c39aff3025...` | Easy care panel with pillow stack and care callouts. | Basketball collage bedding easy care panel | 40 | 30 | 20 | 10 | **100** |
| 229 | 7 | `229_07.jpg` | 47538194055367 | 1000x1000 | `63eb079c5a...` | Size dimension panel showing twin, full, queen and king bed diagrams. | Basketball collage comforter size panel | 40 | 30 | 20 | 10 | **100** |
| 230 | 1 | `230_01.jpg` | 47538169807047 | 1000x1000 | `a5faee1e4a...` | Bedroom mockup of shattered-glass basketball comforter with script sample name Aiden and number 17 on pillows. | Custom shattered glass basketball comforter | 40 | 30 | 20 | 10 | **100** |
| 230 | 2 | `230_02.jpg` | 47538169839815 | 1000x1000 | `9d0e9532cb...` | Second bed mockup of shattered-glass basketball bedding with folded white duvet edge. | Shattered glass basketball bedding mockup | 40 | 30 | 20 | 10 | **100** |
| 230 | 3 | `230_03.jpg` | 47538169872583 | 1000x1000 | `8b3c12074c...` | Duvet cover set versus comforter set comparison panel. | Shattered glass basketball bedding type panel | 40 | 30 | 20 | 10 | **100** |
| 230 | 4 | `230_04.jpg` | 47538169905351 | 1000x1000 | `924ec6eeac...` | Feature panel with basketball print close-ups, microfiber and 3D printed pattern callouts. | Shattered glass basketball comforter feature panel | 40 | 30 | 20 | 10 | **100** |
| 230 | 5 | `230_05.jpg` | 47538169938119 | 1000x1000 | `7cd4ce4e49...` | Low-angle bed mockup of shattered-glass basketball bedding with comfort icon strip. | Shattered glass basketball bedding close view | 40 | 30 | 20 | 10 | **100** |
| 230 | 6 | `230_06.jpg` | 47538169970887 | 1000x1000 | `2f82ab909c...` | Size dimension panel showing twin, full, queen and king bed diagrams. | Shattered glass basketball comforter size panel | 40 | 30 | 20 | 10 | **100** |

---

## 5. Kết quả Kiểm toán Nhu cầu Tìm kiếm & Bằng chứng SERP (20 Truy vấn)

20 truy vấn kiểm chứng độc lập trên Google US SERP và các sàn thương mại điện tử chuyên ngành decor thể thao tại Mỹ:

| ID Truy vấn | Pos | Truy vấn kiểm chứng (Search Query) | Loại | Mục đích tìm kiếm (Buyer Intent) | Các Domain Hàng đầu | Đánh giá Mức độ Phù hợp |
|:---|:---:|:---|:---:|:---|:---|:---|
| `SERP-221-PRI` | 221 | **custom basketball hoop blanket** | `PRIMARY` | Commercial / Custom Basketball Hoop Blanket | etsy.com, zazzle.com, walmart.com | Directly matches Pos 221 glowing rim, net artwork and blanket form. |
| `SERP-221-COM` | 221 | **personalized basketball hoop throw blanket** | `COMPARATOR` | Commercial / Basketball Throw Blanket Gift | etsy.com, amazon.com, personalizationmall.com | Validates buyer preference for throw blanket gift terminology. |
| `SERP-222-PRI` | 222 | **custom basketball shoes blanket** | `PRIMARY` | Commercial / Basketball Shoes Graphic Blanket | etsy.com, amazon.com, zazzle.com | Directly captures Pos 222's distinct ball-beside-shoes graphic. |
| `SERP-222-COM` | 222 | **personalized basketball sneakers throw blanket** | `COMPARATOR` | Commercial / Basketball Sneaker Gift Blanket | etsy.com, redbubble.com, amazon.com | Confirms long-tail variant alignment. |
| `SERP-223-PRI` | 223 | **custom court lines basketball comforter** | `PRIMARY` | Commercial / Court Lines Bedding Set | etsy.com, amazon.com, wayfair.com | Perfect match for Pos 223 gold court lines and cracked court texture. |
| `SERP-223-COM` | 223 | **personalized basketball court bedding** | `COMPARATOR` | Commercial / Basketball Court Bedroom Set | etsy.com, potterybarnkids.com, amazon.com | Supports primary search clustering. |
| `SERP-224-PRI` | 224 | **custom basketball court perspective comforter** | `PRIMARY` | Commercial / 3D Perspective Basketball Bedding | etsy.com, amazon.com, walmart.com | Matches black and red court-perspective bedding design. |
| `SERP-224-COM` | 224 | **personalized basketball comforter set** | `COMPARATOR` | Commercial / Basketball Bedding Set with Pillowcases | amazon.com, etsy.com, kohls.com | Confirms commercial category demand. |
| `SERP-225-PRI` | 225 | **custom hardwood court basketball comforter** | `PRIMARY` | Commercial / Hardwood Arena Floor Bedding | etsy.com, amazon.com, zazzle.com | Accurately targets Pos 225 blue-and-gold hardwood floor artwork. |
| `SERP-225-COM` | 225 | **hardwood basketball court bedding** | `COMPARATOR` | Commercial / Basketball Floor Bedding | amazon.com, walmart.com, wayfair.com | Confirms theme relevance. |
| `SERP-226-PRI` | 226 | **custom neon basketball player blanket** | `PRIMARY` | Commercial / Neon Silhouette Sports Blanket | etsy.com, amazon.com, zazzle.com | Matches Pos 226 neon dunking player artwork. |
| `SERP-226-COM` | 226 | **neon basketball throw blanket** | `COMPARATOR` | Commercial / Neon Sports Throw | amazon.com, etsy.com, walmart.com | Validates neon aesthetic demand. |
| `SERP-227-PRI` | 227 | **custom flaming basketball player blanket** | `PRIMARY` | Commercial / Flaming Basketball Player Blanket | etsy.com, amazon.com, walmart.com | Accurately targets Pos 227 front-facing dribbler with flame effects. |
| `SERP-227-COM` | 227 | **flaming basketball throw blanket** | `COMPARATOR` | Commercial / Fire Sports Throw Blanket | amazon.com, etsy.com, walmart.com | Confirms motif popularity. |
| `SERP-228-PRI` | 228 | **custom dribbling silhouette basketball blanket** | `PRIMARY` | Commercial / Dribbling Silhouette Blanket | etsy.com, amazon.com, zazzle.com | Directly targets Pos 228 black silhouette dribbler and orange graphic. |
| `SERP-228-COM` | 228 | **basketball player silhouette blanket** | `COMPARATOR` | Commercial / Basketball Silhouette Decor | amazon.com, etsy.com, society6.com | Validates silhouette theme demand. |
| `SERP-229-PRI` | 229 | **custom basketball collage comforter** | `PRIMARY` | Commercial / Multi-Element Sports Collage Bedding | etsy.com, amazon.com, wayfair.com | Accurately represents Pos 229 collage artwork. |
| `SERP-229-COM` | 229 | **personalized basketball collage bedding** | `COMPARATOR` | Commercial / Basketball Collage Room Decor | amazon.com, etsy.com, walmart.com | Supports long-tail clustering. |
| `SERP-230-PRI` | 230 | **custom shattered glass basketball comforter** | `PRIMARY` | Commercial / 3D Shattered Effect Basketball Bedding | etsy.com, amazon.com, walmart.com | Directly matches Pos 230 dramatic shattered glass motif. |
| `SERP-230-COM` | 230 | **personalized shattered basketball comforter** | `COMPARATOR` | Commercial / Shattered Basketball Bedding Set | amazon.com, etsy.com, redbubble.com | Validates distinctive motif phrasing. |

---

## 6. Đề xuất Nội dung Thay thế Chuẩn Publish-Ready (Actionable Fixes)

### 6.1. Đề xuất Meta Description Hoàn chỉnh (10 Sản phẩm — Chuẩn 151–160 ký tự, không bị cắt cụt)

| Pos | Đề xuất r2 (Bị cắt cụt ở 155 ký tự) | Độ dài cũ | Đề xuất QA thay thế hoàn chỉnh (Editorial Target 145–165 ký tự) | Độ dài mới |
|:---:|:---|:---:|:---|:---:|
| **221** | `Shop custom basketball hoop blanket with dark blanket artwork with a close-up glowing basketball hoop, white net, sample name and jersey number, selectable` | 155 | **Custom basketball hoop blanket featuring glowing rim and net artwork. Personalize with your player name and jersey number in cozy fleece or plush sherpa.** | **153** |
| **222** | `Shop custom basketball shoes blanket with dark blanket artwork with a basketball beside worn athletic shoes, script sample name and jersey number, selectab` | 155 | **Personalized basketball shoes blanket showcasing court sneakers and game ball graphics. Add custom name and number in ultra-soft fleece or warm sherpa.** | **151** |
| **223** | `Shop custom court lines basketball comforter with black comforter with gold court lines, basketball graphics, cracked texture, sample name and jersey numbe` | 155 | **Custom basketball comforter featuring gold court lines and cracked texture artwork. Personalize with your name and jersey number on soft microfiber bedding.** | **156** |
| **224** | `Shop custom basketball court perspective comforter with black and red court-perspective comforter with a large basketball, sample name and jersey number, s` | 155 | **Custom basketball court perspective comforter in vibrant black and red sports styling. Personalize with name and jersey number on premium microfiber bedding.** | **157** |
| **225** | `Shop custom hardwood court basketball comforter with blue and gold hardwood-court comforter with a large basketball, sample name and jersey number, selecta` | 155 | **Custom hardwood court basketball comforter with blue and gold arena artwork. Personalize with player name and number on ultra-soft polyester microfiber.** | **152** |
| **226** | `Shop custom neon basketball player blanket with dark blanket with neon-style dunking player silhouette, hoop, sample name and jersey number, selectable siz` | 155 | **Custom neon basketball player blanket featuring a vibrant dunking silhouette and hoop graphic. Add your personalized name and number in fleece or sherpa.** | **153** |
| **227** | `Shop custom flaming basketball player blanket with red and navy blanket with a front-facing dribbling player, flame effect, sample name and jersey number, ` | 155 | **Custom flaming basketball player blanket with dynamic dribbler artwork and fire effects. Personalize with your name and number in cozy fleece or sherpa.** | **152** |
| **228** | `Shop custom dribbling silhouette basketball blanket with gray and orange blanket with a black dribbling player silhouette, oversized ball, sample name and ` | 155 | **Custom dribbling silhouette basketball blanket with modern gray and orange graphics. Add player name and number on machine-washable fleece or sherpa throw.** | **155** |
| **229** | `Shop custom basketball collage comforter with black, gray and orange comforter collage with basketballs, hoops, player silhouettes, sample name and jersey ` | 155 | **Custom basketball collage comforter featuring hoops, basketballs, and player silhouettes. Personalize with name and jersey number on soft microfiber bedding.** | **157** |
| **230** | `Shop custom shattered glass basketball comforter with dark shattered-glass comforter with a large basketball, script sample name and jersey number on pillo` | 155 | **Custom shattered glass basketball comforter with 3D cracked artwork. Personalize with name and number; optional matching pillow shams and flat sheet available.** | **159** |

### 6.2. Đề xuất Mô tả Sản phẩm HTML Thay thế Hoàn chỉnh (10 Sản phẩm — Sạch boilerplate, đầy đủ thông số)

#### Sản phẩm Pos 221: Custom Basketball Hoop Blanket
- **Handle:** `personalized-basketball-hoop-with-net-close-up-blanket-with-name-and-number-cc637164fc-cc637164fc`
- **Product Type:** `Blanket`

```html
<p>Custom Basketball Hoop Blanket brings courtside excitement and cozy warmth to your bedroom or living space. Featuring dark blanket artwork featuring a close-up glowing basketball hoop with crisp white net cords, this custom sports throw blanket makes an exceptional personalized gift for basketball players, coaches, and dedicated fans.</p>
<h3>Artwork & Personalization</h3>
<ul>
  <li><strong>Design Motif:</strong> Dark blanket artwork featuring a close-up glowing basketball hoop with crisp white net cords.</li>
  <li><strong>Custom Name Field:</strong> Optional custom name entry (1–200 characters). Enter your recipient's name or leave blank/enter "No" if no name is desired.</li>
  <li><strong>Custom Number Field:</strong> Optional jersey number entry (1–20 characters). Personalize with a favorite team or player number, or enter "No" for unnumbered artwork.</li>
  <li><strong>Mockup Preview Note:</strong> Sample names and numbers shown in gallery mockups are illustrative examples and will be replaced with your exact custom text.</li>
</ul>
<h3>Premium Fabric Options</h3>
<ul>
  <li><strong>Ultra-Soft Fleece:</strong> Lightweight, smooth-faced polyester fleece with vibrant full-color printing; ideal for all-season relaxing, road trips, and game day.</li>
  <li><strong>Plush Sherpa Fleece:</strong> Premium dual-layer construction featuring a smooth printed front paired with a thick, fluffy faux-sherpa lining for maximum warmth.</li>
</ul>
<h3>Available Sizes</h3>
<ul>
  <li><strong>40" x 30" (Baby / Lap):</strong> Compact size for cribs, strollers, toddlers, or lap coverage.</li>
  <li><strong>50" x 40" (Small / Youth):</strong> Perfect throw size for kids, gaming chairs, or travel.</li>
  <li><strong>60" x 50" (Medium / Teen):</strong> Versatile couch throw blanket for teens and movie nights.</li>
  <li><strong>80" x 60" (Large / Adult):</strong> Generous full-coverage blanket for twin/queen beds and lounging.</li>
</ul>
<h3>Care Instructions</h3>
<ul>
  <li>Machine wash cold separately on gentle cycle with mild detergent.</li>
  <li>Tumble dry on low heat or hang air-dry without heat to maintain plush softness.</li>
  <li>Do not bleach, do not iron, and avoid fabric softeners.</li>
</ul>
```

#### Sản phẩm Pos 222: Custom Basketball Shoes Blanket
- **Handle:** `personalized-basketball-next-to-athletic-shoes-blanket-with-name-and-number-dfbf4fde2f-dfbf4fde2f`
- **Product Type:** `Blanket`

```html
<p>Custom Basketball Shoes Blanket brings courtside excitement and cozy warmth to your bedroom or living space. Featuring dark blanket artwork highlighting a basketball resting alongside well-worn athletic basketball shoes, this custom sports throw blanket makes an exceptional personalized gift for basketball players, coaches, and dedicated fans.</p>
<h3>Artwork & Personalization</h3>
<ul>
  <li><strong>Design Motif:</strong> Dark blanket artwork highlighting a basketball resting alongside well-worn athletic basketball shoes.</li>
  <li><strong>Custom Name Field:</strong> Optional custom name entry (1–200 characters). Enter your recipient's name or leave blank/enter "No" if no name is desired.</li>
  <li><strong>Custom Number Field:</strong> Optional jersey number entry (1–20 characters). Personalize with a favorite team or player number, or enter "No" for unnumbered artwork.</li>
  <li><strong>Mockup Preview Note:</strong> Sample names and numbers shown in gallery mockups are illustrative examples and will be replaced with your exact custom text.</li>
</ul>
<h3>Premium Fabric Options</h3>
<ul>
  <li><strong>Ultra-Soft Fleece:</strong> Lightweight, smooth-faced polyester fleece with vibrant full-color printing; ideal for all-season relaxing, road trips, and game day.</li>
  <li><strong>Plush Sherpa Fleece:</strong> Premium dual-layer construction featuring a smooth printed front paired with a thick, fluffy faux-sherpa lining for maximum warmth.</li>
</ul>
<h3>Available Sizes</h3>
<ul>
  <li><strong>40" x 30" (Baby / Lap):</strong> Compact size for cribs, strollers, toddlers, or lap coverage.</li>
  <li><strong>50" x 40" (Small / Youth):</strong> Perfect throw size for kids, gaming chairs, or travel.</li>
  <li><strong>60" x 50" (Medium / Teen):</strong> Versatile couch throw blanket for teens and movie nights.</li>
  <li><strong>80" x 60" (Large / Adult):</strong> Generous full-coverage blanket for twin/queen beds and lounging.</li>
</ul>
<h3>Care Instructions</h3>
<ul>
  <li>Machine wash cold separately on gentle cycle with mild detergent.</li>
  <li>Tumble dry on low heat or hang air-dry without heat to maintain plush softness.</li>
  <li>Do not bleach, do not iron, and avoid fabric softeners.</li>
</ul>
```

#### Sản phẩm Pos 223: Custom Court Lines Basketball Comforter
- **Handle:** `personalized-basketball-on-court-lines-comforter-name-number-303fbd50f6-303fbd50f6`
- **Product Type:** `Comforter`

```html
<p>Custom Court Lines Basketball Comforter transforms any bedroom into the ultimate sports retreat. Featuring bold black and gold court lines with authentic cracked arena court texture and basketball graphic elements, this personalized bedding ensemble delivers bold courtside style and all-season sleeping comfort.</p>
<h3>Design & Customization</h3>
<ul>
  <li><strong>Artwork Motif:</strong> Bold black and gold court lines with authentic cracked arena court texture and basketball graphic elements. High-definition 3D digital print on ultra-soft polyester microfiber.</li>
  <li><strong>Customize Your Name:</strong> Required personalization field (1–30 characters). Enter the player's name or enter "NO" if you prefer uncustomized bedding.</li>
  <li><strong>Customize Your Number:</strong> Required jersey number field (1–5 characters, e.g., "10" or "23"). Enter "NO" for a clean graphic without numbers.</li>
  <li><strong>Sample Display Note:</strong> Names and numbers displayed in gallery mockups demonstrate custom layout placement and will be replaced with your submission.</li>
</ul>
<h3>Bedding Type & Configuration Options</h3>
<ul>
  <li><strong>Duvet Cover:</strong> Lightweight protective cover equipped with a concealed bottom zipper closure and interior corner ties to secure your duvet insert (insert not included).</li>
  <li><strong>Comforter:</strong> Ready-to-use plush bedspread filled with lightweight down-alternative microfiber batting for cozy, year-round warmth.</li>
  <li><strong>Available Bed Sizes:</strong> Twin (68" x 88"), Full (78" x 88"), Queen (88" x 88"), and King (104" x 88").</li>
  <li><strong>Optional Pillowcases:</strong> Matching pillowcases are optional add-ons selectable via 'Choose Pillowcases' (None, 1, or 2 pillowcases) and are not included when 'None' is selected.</li>
  <li><strong>Optional Sheet Cover:</strong> Add an optional flat sheet matching your selected bedding dimensions via 'Additional Sheet Cover (Flat Sheet)'.</li>
</ul>
<h3>Material & Care</h3>
<ul>
  <li>100% premium brushed polyester microfiber; breathable, hypoallergenic, and wrinkle-resistant.</li>
  <li>Machine wash cold on gentle cycle; tumble dry low or air dry. Do not bleach or dry clean.</li>
</ul>
```

#### Sản phẩm Pos 224: Custom Basketball Court Perspective Comforter
- **Handle:** `personalized-basketball-on-court-perspective-comforter-with-name-and-number-1acfbe4260-1acfbe4260`
- **Product Type:** `Comforter`

```html
<p>Custom Basketball Court Perspective Comforter transforms any bedroom into the ultimate sports retreat. Featuring dramatic perspective-angle court layout in black and red sports tones with a prominent basketball graphic, this personalized bedding ensemble delivers bold courtside style and all-season sleeping comfort.</p>
<h3>Design & Customization</h3>
<ul>
  <li><strong>Artwork Motif:</strong> Dramatic perspective-angle court layout in black and red sports tones with a prominent basketball graphic. High-definition 3D digital print on ultra-soft polyester microfiber.</li>
  <li><strong>Customize Your Name:</strong> Required personalization field (1–30 characters). Enter the player's name or enter "NO" if you prefer uncustomized bedding.</li>
  <li><strong>Customize Your Number:</strong> Required jersey number field (1–5 characters, e.g., "10" or "23"). Enter "NO" for a clean graphic without numbers.</li>
  <li><strong>Sample Display Note:</strong> Names and numbers displayed in gallery mockups demonstrate custom layout placement and will be replaced with your submission.</li>
</ul>
<h3>Bedding Type & Configuration Options</h3>
<ul>
  <li><strong>Duvet Cover:</strong> Lightweight protective cover equipped with a concealed bottom zipper closure and interior corner ties to secure your duvet insert (insert not included).</li>
  <li><strong>Comforter:</strong> Ready-to-use plush bedspread filled with lightweight down-alternative microfiber batting for cozy, year-round warmth.</li>
  <li><strong>Available Bed Sizes:</strong> Twin (68" x 88"), Full (78" x 88"), Queen (88" x 88"), and King (104" x 88").</li>
  <li><strong>Optional Pillowcases:</strong> Matching pillowcases are optional add-ons selectable via 'Choose Pillowcases' (None, 1, or 2 pillowcases) and are not included when 'None' is selected.</li>
  <li><strong>Optional Sheet Cover:</strong> Add an optional flat sheet matching your selected bedding dimensions via 'Additional Sheet Cover (Flat Sheet)'.</li>
</ul>
<h3>Material & Care</h3>
<ul>
  <li>100% premium brushed polyester microfiber; breathable, hypoallergenic, and wrinkle-resistant.</li>
  <li>Machine wash cold on gentle cycle; tumble dry low or air dry. Do not bleach or dry clean.</li>
</ul>
```

#### Sản phẩm Pos 225: Custom Hardwood Court Basketball Comforter
- **Handle:** `personalized-basketball-on-hardwood-court-comforter-with-name-and-number-8aee799947-8aee799947`
- **Product Type:** `Comforter`

```html
<p>Custom Hardwood Court Basketball Comforter transforms any bedroom into the ultimate sports retreat. Featuring blue and gold arena court artwork with a realistic hardwood floor printed texture and large basketball, this personalized bedding ensemble delivers bold courtside style and all-season sleeping comfort.</p>
<h3>Design & Customization</h3>
<ul>
  <li><strong>Artwork Motif:</strong> Blue and gold arena court artwork with a realistic hardwood floor printed texture and large basketball. High-definition 3D digital print on ultra-soft polyester microfiber.</li>
  <li><strong>Customize Your Name:</strong> Required personalization field (1–30 characters). Enter the player's name or enter "NO" if you prefer uncustomized bedding.</li>
  <li><strong>Customize Your Number:</strong> Required jersey number field (1–5 characters, e.g., "10" or "23"). Enter "NO" for a clean graphic without numbers.</li>
  <li><strong>Sample Display Note:</strong> Names and numbers displayed in gallery mockups demonstrate custom layout placement and will be replaced with your submission.</li>
</ul>
<h3>Bedding Type & Configuration Options</h3>
<ul>
  <li><strong>Duvet Cover:</strong> Lightweight protective cover equipped with a concealed bottom zipper closure and interior corner ties to secure your duvet insert (insert not included).</li>
  <li><strong>Comforter:</strong> Ready-to-use plush bedspread filled with lightweight down-alternative microfiber batting for cozy, year-round warmth.</li>
  <li><strong>Available Bed Sizes:</strong> Twin (68" x 88"), Full (78" x 88"), Queen (88" x 88"), and King (104" x 88").</li>
  <li><strong>Optional Pillowcases:</strong> Matching pillowcases are optional add-ons selectable via 'Choose Pillowcases' (None, 1, or 2 pillowcases) and are not included when 'None' is selected.</li>
  <li><strong>Optional Sheet Cover:</strong> Add an optional flat sheet matching your selected bedding dimensions via 'Additional Sheet Cover (Flat Sheet)'.</li>
</ul>
<h3>Material & Care</h3>
<ul>
  <li>100% premium brushed polyester microfiber; breathable, hypoallergenic, and wrinkle-resistant.</li>
  <li>Machine wash cold on gentle cycle; tumble dry low or air dry. Do not bleach or dry clean.</li>
</ul>
```

#### Sản phẩm Pos 226: Custom Neon Basketball Player Blanket
- **Handle:** `personalized-neon-basketball-player-blanket-788fb6ae63-788fb6ae63`
- **Product Type:** `Blanket`

```html
<p>Custom Neon Basketball Player Blanket brings courtside excitement and cozy warmth to your bedroom or living space. Featuring electrifying neon-style player silhouette frozen in a powerful slam dunk beside a neon basketball hoop, this custom sports throw blanket makes an exceptional personalized gift for basketball players, coaches, and dedicated fans.</p>
<h3>Artwork & Personalization</h3>
<ul>
  <li><strong>Design Motif:</strong> Electrifying neon-style player silhouette frozen in a powerful slam dunk beside a neon basketball hoop.</li>
  <li><strong>Custom Name Field:</strong> Optional custom name entry (1–200 characters). Enter your recipient's name or leave blank/enter "No" if no name is desired.</li>
  <li><strong>Custom Number Field:</strong> Optional jersey number entry (1–20 characters). Personalize with a favorite team or player number, or enter "No" for unnumbered artwork.</li>
  <li><strong>Mockup Preview Note:</strong> Sample names and numbers shown in gallery mockups are illustrative examples and will be replaced with your exact custom text.</li>
</ul>
<h3>Premium Fabric Options</h3>
<ul>
  <li><strong>Ultra-Soft Fleece:</strong> Lightweight, smooth-faced polyester fleece with vibrant full-color printing; ideal for all-season relaxing, road trips, and game day.</li>
  <li><strong>Plush Sherpa Fleece:</strong> Premium dual-layer construction featuring a smooth printed front paired with a thick, fluffy faux-sherpa lining for maximum warmth.</li>
</ul>
<h3>Available Sizes</h3>
<ul>
  <li><strong>40" x 30" (Baby / Lap):</strong> Compact size for cribs, strollers, toddlers, or lap coverage.</li>
  <li><strong>50" x 40" (Small / Youth):</strong> Perfect throw size for kids, gaming chairs, or travel.</li>
  <li><strong>60" x 50" (Medium / Teen):</strong> Versatile couch throw blanket for teens and movie nights.</li>
  <li><strong>80" x 60" (Large / Adult):</strong> Generous full-coverage blanket for twin/queen beds and lounging.</li>
</ul>
<h3>Care Instructions</h3>
<ul>
  <li>Machine wash cold separately on gentle cycle with mild detergent.</li>
  <li>Tumble dry on low heat or hang air-dry without heat to maintain plush softness.</li>
  <li>Do not bleach, do not iron, and avoid fabric softeners.</li>
</ul>
```

#### Sản phẩm Pos 227: Custom Flaming Basketball Player Blanket
- **Handle:** `personalized-basketball-player-front-blanket-with-name-and-number-258e7f9e2e-258e7f9e2e`
- **Product Type:** `Blanket`

```html
<p>Custom Flaming Basketball Player Blanket brings courtside excitement and cozy warmth to your bedroom or living space. Featuring action-packed front-facing dribbling basketball player engulfed in vibrant flame and fire effects, this custom sports throw blanket makes an exceptional personalized gift for basketball players, coaches, and dedicated fans.</p>
<h3>Artwork & Personalization</h3>
<ul>
  <li><strong>Design Motif:</strong> Action-packed front-facing dribbling basketball player engulfed in vibrant flame and fire effects.</li>
  <li><strong>Custom Name Field:</strong> Optional custom name entry (1–200 characters). Enter your recipient's name or leave blank/enter "No" if no name is desired.</li>
  <li><strong>Custom Number Field:</strong> Optional jersey number entry (1–20 characters). Personalize with a favorite team or player number, or enter "No" for unnumbered artwork.</li>
  <li><strong>Mockup Preview Note:</strong> Sample names and numbers shown in gallery mockups are illustrative examples and will be replaced with your exact custom text.</li>
</ul>
<h3>Premium Fabric Options</h3>
<ul>
  <li><strong>Ultra-Soft Fleece:</strong> Lightweight, smooth-faced polyester fleece with vibrant full-color printing; ideal for all-season relaxing, road trips, and game day.</li>
  <li><strong>Plush Sherpa Fleece:</strong> Premium dual-layer construction featuring a smooth printed front paired with a thick, fluffy faux-sherpa lining for maximum warmth.</li>
</ul>
<h3>Available Sizes</h3>
<ul>
  <li><strong>40" x 30" (Baby / Lap):</strong> Compact size for cribs, strollers, toddlers, or lap coverage.</li>
  <li><strong>50" x 40" (Small / Youth):</strong> Perfect throw size for kids, gaming chairs, or travel.</li>
  <li><strong>60" x 50" (Medium / Teen):</strong> Versatile couch throw blanket for teens and movie nights.</li>
  <li><strong>80" x 60" (Large / Adult):</strong> Generous full-coverage blanket for twin/queen beds and lounging.</li>
</ul>
<h3>Care Instructions</h3>
<ul>
  <li>Machine wash cold separately on gentle cycle with mild detergent.</li>
  <li>Tumble dry on low heat or hang air-dry without heat to maintain plush softness.</li>
  <li>Do not bleach, do not iron, and avoid fabric softeners.</li>
</ul>
```

#### Sản phẩm Pos 228: Custom Dribbling Silhouette Basketball Blanket
- **Handle:** `personalized-basketball-player-silhouette-blanket-with-name-and-number-35f8822fba-35f8822fba`
- **Product Type:** `Blanket`

```html
<p>Custom Dribbling Silhouette Basketball Blanket brings courtside excitement and cozy warmth to your bedroom or living space. Featuring modern minimalist black dribbler silhouette against sleek gray and orange court graphics with an oversized ball, this custom sports throw blanket makes an exceptional personalized gift for basketball players, coaches, and dedicated fans.</p>
<h3>Artwork & Personalization</h3>
<ul>
  <li><strong>Design Motif:</strong> Modern minimalist black dribbler silhouette against sleek gray and orange court graphics with an oversized ball.</li>
  <li><strong>Custom Name Field:</strong> Optional custom name entry (1–200 characters). Enter your recipient's name or leave blank/enter "No" if no name is desired.</li>
  <li><strong>Custom Number Field:</strong> Optional jersey number entry (1–20 characters). Personalize with a favorite team or player number, or enter "No" for unnumbered artwork.</li>
  <li><strong>Mockup Preview Note:</strong> Sample names and numbers shown in gallery mockups are illustrative examples and will be replaced with your exact custom text.</li>
</ul>
<h3>Premium Fabric Options</h3>
<ul>
  <li><strong>Ultra-Soft Fleece:</strong> Lightweight, smooth-faced polyester fleece with vibrant full-color printing; ideal for all-season relaxing, road trips, and game day.</li>
  <li><strong>Plush Sherpa Fleece:</strong> Premium dual-layer construction featuring a smooth printed front paired with a thick, fluffy faux-sherpa lining for maximum warmth.</li>
</ul>
<h3>Available Sizes</h3>
<ul>
  <li><strong>40" x 30" (Baby / Lap):</strong> Compact size for cribs, strollers, toddlers, or lap coverage.</li>
  <li><strong>50" x 40" (Small / Youth):</strong> Perfect throw size for kids, gaming chairs, or travel.</li>
  <li><strong>60" x 50" (Medium / Teen):</strong> Versatile couch throw blanket for teens and movie nights.</li>
  <li><strong>80" x 60" (Large / Adult):</strong> Generous full-coverage blanket for twin/queen beds and lounging.</li>
</ul>
<h3>Care Instructions</h3>
<ul>
  <li>Machine wash cold separately on gentle cycle with mild detergent.</li>
  <li>Tumble dry on low heat or hang air-dry without heat to maintain plush softness.</li>
  <li>Do not bleach, do not iron, and avoid fabric softeners.</li>
</ul>
```

#### Sản phẩm Pos 229: Custom Basketball Collage Comforter
- **Handle:** `personalized-basketball-players-and-hoops-comforter-with-name-and-number-ff13bad888-ff13bad888`
- **Product Type:** `Comforter`

```html
<p>Custom Basketball Collage Comforter transforms any bedroom into the ultimate sports retreat. Featuring dynamic multi-panel basketball collage showcasing basketballs, hoops, backboards, and athletic player silhouettes, this personalized bedding ensemble delivers bold courtside style and all-season sleeping comfort.</p>
<h3>Design & Customization</h3>
<ul>
  <li><strong>Artwork Motif:</strong> Dynamic multi-panel basketball collage showcasing basketballs, hoops, backboards, and athletic player silhouettes. High-definition 3D digital print on ultra-soft polyester microfiber.</li>
  <li><strong>Customize Your Name:</strong> Required personalization field (1–30 characters). Enter the player's name or enter "NO" if you prefer uncustomized bedding.</li>
  <li><strong>Customize Your Number:</strong> Required jersey number field (1–5 characters, e.g., "10" or "23"). Enter "NO" for a clean graphic without numbers.</li>
  <li><strong>Sample Display Note:</strong> Names and numbers displayed in gallery mockups demonstrate custom layout placement and will be replaced with your submission.</li>
</ul>
<h3>Bedding Type & Configuration Options</h3>
<ul>
  <li><strong>Duvet Cover:</strong> Lightweight protective cover equipped with a concealed bottom zipper closure and interior corner ties to secure your duvet insert (insert not included).</li>
  <li><strong>Comforter:</strong> Ready-to-use plush bedspread filled with lightweight down-alternative microfiber batting for cozy, year-round warmth.</li>
  <li><strong>Available Bed Sizes:</strong> Twin (68" x 88"), Full (78" x 88"), Queen (88" x 88"), and King (104" x 88").</li>
  <li><strong>Optional Pillowcases:</strong> Matching pillowcases are optional add-ons selectable via 'Choose Pillowcases' (None, 1, or 2 pillowcases) and are not included when 'None' is selected.</li>
  <li><strong>Optional Sheet Cover:</strong> Add an optional flat sheet matching your selected bedding dimensions via 'Additional Sheet Cover (Flat Sheet)'.</li>
</ul>
<h3>Material & Care</h3>
<ul>
  <li>100% premium brushed polyester microfiber; breathable, hypoallergenic, and wrinkle-resistant.</li>
  <li>Machine wash cold on gentle cycle; tumble dry low or air dry. Do not bleach or dry clean.</li>
</ul>
```

#### Sản phẩm Pos 230: Custom Shattered Glass Basketball Comforter
- **Handle:** `personalized-basketball-shattered-glass-comforter-with-name-and-number-353fc0c17e-353fc0c17e`
- **Product Type:** `Comforter`

```html
<p>Custom Shattered Glass Basketball Comforter transforms any bedroom into the ultimate sports retreat. Featuring dramatic 3D shattered-glass explosion artwork centered around a basketball bursting through broken glass, this personalized bedding ensemble delivers bold courtside style and all-season sleeping comfort.</p>
<h3>Design & Customization</h3>
<ul>
  <li><strong>Artwork Motif:</strong> Dramatic 3d shattered-glass explosion artwork centered around a basketball bursting through broken glass. High-definition 3D digital print on ultra-soft polyester microfiber.</li>
  <li><strong>Customize Your Name:</strong> Required personalization field (1–30 characters). Enter the player's name or enter "NO" if you prefer uncustomized bedding.</li>
  <li><strong>Customize Your Number:</strong> Required jersey number field (1–5 characters, e.g., "10" or "23"). Enter "NO" for a clean graphic without numbers.</li>
  <li><strong>Sample Display Note:</strong> Names and numbers displayed in gallery mockups demonstrate custom layout placement and will be replaced with your submission.</li>
</ul>
<h3>Bedding Type & Configuration Options</h3>
<ul>
  <li><strong>Duvet Cover:</strong> Lightweight protective cover equipped with a concealed bottom zipper closure and interior corner ties to secure your duvet insert (insert not included).</li>
  <li><strong>Comforter:</strong> Ready-to-use plush bedspread filled with lightweight down-alternative microfiber batting for cozy, year-round warmth.</li>
  <li><strong>Available Bed Sizes:</strong> Twin (68" x 88"), Full (78" x 88"), Queen (88" x 88"), and King (104" x 88").</li>
  <li><strong>Optional Pillow Shams:</strong> Pillow shams with matching artwork and custom personalization are optional add-ons available via the 'Choose Pillowcases' menu (None, 1, or 2 pillowcases). Pillow shams are not included by default.</li>
  <li><strong>Optional Sheet Cover:</strong> Add an optional flat sheet matching your selected bedding dimensions via 'Additional Sheet Cover (Flat Sheet)'.</li>
</ul>
<h3>Material & Care</h3>
<ul>
  <li>100% premium brushed polyester microfiber; breathable, hypoallergenic, and wrinkle-resistant.</li>
  <li>Machine wash cold on gentle cycle; tumble dry low or air dry. Do not bleach or dry clean.</li>
</ul>
```

---

## 7. Kết luận Kiểm định & Bàn giao

- **Kết luận chung lô Batch 23 revision r2:** `QA_REVISE`.
- **Lý do:** Điểm số toàn lô đạt mức rất cao **95.0/100**, không có lỗi vi phạm nghiêm trọng (0 CRITICAL), tuy nhiên tồn tại 10 lỗi MAJOR tại tiêu chí D2 (chứa prompt boilerplate và mâu thuẫn tùy chọn blanket) cùng 10 lỗi MINOR tại tiêu chí D1 (cắt cụt 155 ký tự).
- **Lưu ý nghiệp vụ:** Kết quả `QA_REVISE` là báo cáo kiểm định kỹ thuật độc lập; **không cấu thành phê duyệt (APPROVED) và không cấp quyền xuất file import Shopify** cho đến khi các nội dung sửa đổi ở Mục 6 được áp dụng vào bản r3.
- **Tình trạng tiến trình:** Đã hoàn thành toàn bộ kiểm định cho Batch 23 r2. Đặt `awaiting_confirmation=true` và tạm dừng trước khi chuyển sang batch tiếp theo.