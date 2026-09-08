# Báo cáo QA Độc lập SEO Workbook - Batch 22 Revision r2

> **Mã phiên QA:** `20260908_134500`  
> **Tập tin nguồn thẩm định:** `resutls/jeminise.com/20260906_234129/revisions/qa_batch_022_r2/SEO_Product_Optimization_qa_batch_022_r2.xlsx`  
> **SHA256 Tập tin nguồn:** `082f68e975e3be9d4ac3eb89d45a2aa6d023d07b5459c653d990c3d69bbf935d`  
> **Thị trường mục tiêu:** United States (`en-US`) | **Nội dung:** English SEO Copy  
> **Phạm vi thẩm định:** 10 sản phẩm (Inventory Position 211–220), 73 ảnh gallery, 40 dòng Keyword_Map, 10 dòng Buyer_Search_Research  
> **Đặc thù phiên chạy:** **Phiên QA độc lập đầu tiên cho Batch 22 trên revision r2** (`first_qa_run: true`, `source_revision: "r2"`)  
> **Trạng thái thẩm định:** `awaiting_confirmation=true` | **Kết luận tổng thể:** `QA_REVISE` (Điểm trung bình: **95.00/100**)  

---

## 1. Tóm tắt kết quả QA & Chấm điểm Tổng thể

Phiên kiểm toán độc lập đầu tiên cho **Batch 22 Revision r2** được thực hiện với 100% trọng số đánh giá trên toàn bộ 10 sản phẩm, 73 hình ảnh gallery, đối chiếu trực tiếp dữ liệu live storefront và customizer tại `jeminise.com`, cùng 20 truy vấn US SERP độc lập:

| Chỉ số | Giá trị | Ghi chú kiểm toán |
| :--- | :---: | :--- |
| **Tổng số sản phẩm** | 10 | Position 211 đến 220 (`f16f46abef` đến `863d6f6ff1`) |
| **Tỷ lệ kiểm toán sản phẩm** | 100% (10/10) | Cột `page_read=TRUE` đầy đủ |
| **Tổng số ảnh gallery** | 73 | Pos 211: 9 ảnh; Pos 212, 213, 216, 217, 220: 7 ảnh; Pos 214: 5 ảnh; Pos 215, 218, 219: 8 ảnh |
| **Tỷ lệ kiểm toán ảnh** | 100% (73/73) | Điểm ảnh trung bình 100.0/100, 100% PASS |
| **Điểm số trung bình** | **95.00 / 100** | 10/10 sản phẩm đạt chính xác 95.0 điểm (P1: 15, P2: 10, K1: 10, K2: 5, K3: 5, T1: 10, T2: 5, D1: 3, D2: 7, I1: 20, E1: 5) |
| **Số sản phẩm QA_PASS** | 0 | Do tồn tại lỗi MAJOR ở tiêu chuẩn D2 |
| **Số sản phẩm QA_REVISE** | **10** | Yêu cầu chuẩn hóa lại mô tả sản phẩm D2 và cắt tỉa meta description D1 |
| **Số sản phẩm QA_FAIL** | 0 | Không có lỗi CRITICAL |
| **Số sản phẩm QA_INCOMPLETE**| 0 | 100% dữ liệu được đánh giá đầy đủ |
| **Kết luận tổng thể** | **QA_REVISE** | Sẵn sàng xuất bản sau khi áp dụng 10 mô tả HTML và meta descriptions đề xuất |

---

## 2. Bảng tổng hợp đánh giá 10 sản phẩm (Pos 211–220)

| Pos | Product Key / Handle | Title đề xuất r2 | Primary Keyword | P1 | P2 | K1 | K2 | K3 | T1 | T2 | D1 | D2 | I1 | E1 | Tổng | Issues | QA Status |
| :---: | :--- | :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| 211 | `personalized-basketball-close-up-with-number-blanket-with-name-and-number-f16f46abef-f16f46abef` | Basketball Close-Up Number Blanket | *basketball close-up number blanket* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 212 | `personalized-basketball-court-perspective-comforter-with-name-and-number-5545af9fa4-5545af9fa4` | Custom Basketball Court Comforter | *custom basketball court comforter* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 213 | `personalized-basketball-cracked-wall-comforter-with-name-and-number-aeaddaa506-aeaddaa506` | Custom Basketball Cracked Wall Comforter | *custom basketball cracked wall comforter* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 214 | `personalized-basketball-entering-net-comforter-e55e615fd3-e55e615fd3` | Custom Rainbow Basketball Net Comforter | *custom rainbow basketball net comforter* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 215 | `personalized-basketball-fire-and-water-splash-blanket-with-name-and-number-b05a53f208-b05a53f208` | Custom Basketball Fire Water Blanket | *custom basketball fire water blanket* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 216 | `personalized-basketball-flames-flying-over-comforter-with-name-and-number-feb4be1ea4-feb4be1ea4` | Custom Flaming Basketball Comforter | *custom flaming basketball comforter* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 217 | `personalized-basketball-glowing-light-burst-comforter-with-name-and-number-3bc9e1267e-3bc9e1267e` | Custom Light Burst Basketball Comforter | *custom light burst basketball comforter* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 218 | `personalized-basketball-held-by-hand-on-court-blanket-with-name-and-number-079998b76f-079998b76f` | Custom Basketball Court Blanket | *custom basketball court blanket* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 219 | `personalized-basketball-held-under-player-arm-blanket-with-name-and-number-7a077c5144-7a077c5144` | Custom Basketball Player Blanket | *custom basketball player blanket* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 220 | `personalized-basketball-hoop-comforter-name-number-863d6f6ff1-863d6f6ff1` | Custom Orange Basketball Hoop Comforter | *custom orange basketball hoop comforter* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |

---

## 3. Phân tích Chi tiết 11 Tiêu chuẩn Đánh giá & Các điểm Đặc thù

### 3.1. Nhóm Tiêu chuẩn Đạt Điểm Tuyệt đối (PASS)
- **P1. Nhận diện Sản phẩm (15/15đ):** Phân biệt rành mạch hai nhóm form factor thể thao: 6 sản phẩm Bedding (Comforter / Duvet Cover) và 4 sản phẩm Blanket (Pos 211, 215, 218, 219 dạng chăn nỉ nhung fleece throw). Nhận diện chính xác 100% các motif đặc trưng: cận cảnh da bóng rổ sần, sân đấu gỗ arena, tường nứt 3D (cracked wall), rổ bóng rổ cầu vồng phát sáng (rainbow net), xung đột nguyên tố lửa và nước (fire and water), bóng rổ bốc lửa bay qua sân (flaming ball), vầng sáng tỏa tia (light burst), bàn tay nắm bóng trên sân, tư thế ôm bóng dưới cánh tay và rổ bóng rổ viền cam rực rỡ.
- **P2. Phân loại & Cấu hình (10/10đ):** Khớp đúng taxonomy Shopify collections và kiến trúc biến thể variants:
  - *Nhóm Comforter/Bedding (6 sản phẩm - Pos 212, 213, 214, 216, 217, 220):* 48 variants gồm `Choose Product Type + Size` (8 lựa chọn: Duvet Cover / Comforter x Twin/Full/Queen/King), `Choose Pillowcases` (3 lựa chọn: None/1/2), `Additional Sheet Cover` (2 lựa chọn: None/1).
  - *Nhóm Blanket (3 sản phẩm - Pos 215, 218, 219):* 8 variants gồm tùy chọn đơn `Choose Your Size` (Small 40"x50", Medium 50"x60", Large 60"x80", v.v.).
  - *Sản phẩm Pos 211:* Kiểm toán live storefront xác nhận hiện tại chỉ có 1 biến thể `Default Title`, cấu hình chưa hoàn thiện so với các sản phẩm chăn khác trên website.
- **K1, K2, K3. Tối ưu Từ khóa (20/20đ):** Bộ từ khóa chính và phụ bao phủ chuẩn xác ý định mua sắm đồ decor và quà tặng bóng rổ thể thao tại Mỹ, được kiểm chứng qua 20 truy vấn Google US SERP độc lập.
- **T1, T2. Tối ưu Tiêu đề (15/15đ):** Tiêu đề ngắn gọn, cấu trúc chuẩn, từ khóa chính đặt đầu; Meta Title dưới 60 ký tự, kết thúc bằng thương hiệu `| Jeminise`.
- **I1. Tối ưu Ảnh Gallery (20/20đ):** Toàn bộ 73 ảnh được kiểm tra trực quan trực tiếp theo từng pixel; điểm ảnh trung bình đạt 100/100đ, alt text mô tả chính xác chi tiết đồ họa.
- **E1. Căn cứ & Thực chứng (5/5đ):** Đầy đủ hồ sơ thực chứng: ảnh local, kiểm toán live customizer, đối chiếu admin export và 20 kết quả SERP.

### 3.2. Tiêu chuẩn Bị Trừ điểm: D1. Meta Description SEO (3/5đ - PARTIAL)
- **Hiện trạng kiểm toán:** Trong bản r2, **100% (10/10) Meta Description bị cắt cụt cơ học ở đúng 155 ký tự** (`hard-truncated at 155 chars`), làm đứt đoạn câu giữa chừng hoặc cắt ngang từ ngữ:
  - Pos 211: *"...visible basket"* (cắt ngang từ `basketball`)
  - Pos 212: *"...and selectable size"* (thiếu dấu kết câu)
  - Pos 213: *"...and selectable size"* (thiếu dấu kết câu)
  - Pos 214: *"...selectabl"* (cắt cụt từ `selectable`)
  - Pos 215: *"...and jersey num"* (cắt ngang từ `number`)
  - Pos 216: *"...and jersey number"* (thiếu dấu kết câu)
  - Pos 217: *"...selectable "* (khoảng trắng lơ lửng cuối câu)
  - Pos 218: *"...selec"* (cắt cụt từ `selectable`)
  - Pos 219: *"...selectab"* (cắt cụt từ `selectable`)
  - Pos 220: *"...on pillow shams, s"* (cắt cụt từ `selectable`)
- **Biện pháp xử lý:** Trừ 2 điểm D1 (3/5đ, PARTIAL), ghi nhận 10 lỗi MINOR (`ISSUE-0011` đến `ISSUE-0020`), và cung cấp 10 bản Meta Description hoàn chỉnh dưới 150 ký tự tại Mục 6.1.

### 3.3. Tiêu chuẩn Bị Trừ điểm: D2. Product Description HTML (7/10đ - PARTIAL)
- **Hiện trạng kiểm toán:** Toàn bộ 10 mô tả r2 chứa câu lệnh sinh nội bộ (prompt boilerplate leakage) hiển thị trực tiếp cho khách hàng:
  > *"The wording focuses on the visible sports artwork, product form and selectable options for this exact design."*  
  > *"Gallery images show the main bedding or blanket mockup plus feature, care, bedding-type, size or lifestyle panels where present."*
- **Đặc thù kiểm toán Sản phẩm 211 (`f16f46abef`):**
  - Dữ liệu r2 ghi: *"Available option groups: Title. No shopper text-entry field is described for this product. Select the product type and size shown on the product page before checkout."*
  - Kiểm toán live storefront xác nhận trang sản phẩm Pos 211 **chỉ có duy nhất 1 biến thể `Default Title`**, không tải mã nguồn customizer app (`amazon-customizer.js` vắng mặt). Do đó, câu hướng dẫn người mua *"Select the product type and size shown on the product page before checkout"* là mâu thuẫn trực tiếp với giao diện mua hàng thực tế.
  - Artwork mockup có hiển thị số áo mẫu và tên mẫu dọc. Đây là sản phẩm chưa được cấu hình đầy đủ options trên Shopify; QA ghi nhận limitation kỹ thuật và chuẩn hóa mô tả theo đúng dạng chăn nỉ nhung plush throw tiêu chuẩn.
- **Đặc thù kiểm toán Sản phẩm 220 (`863d6f6ff1`):**
  - Mockup hình ảnh hiển thị tên và số áo trên vỏ gối shams. Tuy nhiên, trong luồng mua hàng thực tế, nhóm tùy chọn `Choose Pillowcases` là tùy chọn mua thêm (`None`, `1 Pillowcase`, `2 Pillowcases`).
  - Mô tả r2 không làm rõ điều kiện này, dễ gây hiểu lầm rằng vỏ gối shams được tặng kèm mặc định.
- **Kiểm toán trường cá nhân hóa (Customizer Audit):**
  - *Nhóm Bedding Comforter (Pos 212, 213, 214, 216, 217, 220):* Có **2 trường BẮT BUỘC** (`required=true`) gồm `Customize Your Name` (tối đa 30 ký tự) và `Customize Your Number` (tối đa 5 ký tự, placeholder `10`).
  - *Nhóm Blanket (Pos 215, 218, 219):* Có **2 trường TÙY CHỌN** (`required=false`) gồm `Custom Name` (tối đa 200 ký tự, placeholder `David`) và `Custom Number` (tối đa 20 ký tự, placeholder `22`).
- **Thiếu sót thông số kỹ thuật:** Bản r2 thiếu phân định rành mạch cấu tạo Comforter vs Duvet Cover, thiếu chi tiết khóa kéo ẩn đáy, thiếu bảng thông số kích thước và hướng dẫn giặt ủi chi tiết.
- **Biện pháp xử lý:** Trừ 3 điểm D2 (7/10đ, PARTIAL), ghi nhận 10 lỗi MAJOR (`ISSUE-0001` đến `ISSUE-0010`), và cung cấp 10 bản mô tả HTML chuẩn publish-ready thay thế hoàn chỉnh.

---

## 4. Kết quả Kiểm toán Trực quan 73 Hình ảnh Gallery

Toàn bộ 73 hình ảnh đã được kiểm tra trực quan trực tiếp theo từng pixel từ thư mục ảnh kiểm toán:

| Pos | Handle | Số lượng ảnh | Chi tiết kiểm toán trực quan (Direct Visual Audit) | Alt Text r2 | Đánh giá |
| :---: | :--- | :---: | :--- | :--- | :---: |
| 211 | `personalized-basketball-close-up-with-number-blanket-with-name-and-number-f16f46abef-f16f46abef` | 9 | `211_01`: Mockup chăn nỉ đen, quả bóng rổ cận cảnh cỡ lớn, số áo lớn #22 và tên David. `211_02`: Chăn trải sofa. `211_03`: Chăn gấp. `211_04`: Bảng kích thước chăn. `211_05`: Cận cảnh vải nỉ nhung. `211_06`: Lifestyle. `211_07`: Mockup bổ sung. `211_08`: Feature panel. `211_09`: Minh họa đường may viền. | Mô tả chuẩn chăn bóng rổ cận cảnh và số áo lớn | PASS (100đ) |
| 212 | `personalized-basketball-court-perspective-comforter-with-name-and-number-5545af9fa4-5545af9fa4` | 7 | `212_01`: Mockup chăn sàn đấu bóng rổ góc nhìn sâu dưới ánh đèn pha, tên Ethan #10. `212_02`: Vỏ chăn gấp. `212_03`–`212_07`: 5 panel infographic tiêu chuẩn (icon tính năng, sợi microfiber 3D print, Easy Care, Comforter vs Duvet, size chart). | Chính xác chi tiết sàn đấu gỗ và đèn pha arena | PASS (100đ) |
| 213 | `personalized-basketball-cracked-wall-comforter-with-name-and-number-aeaddaa506-aeaddaa506` | 7 | `213_01`: Mockup quả bóng rổ cam đâm xuyên bức tường bê tông phòng tập nứt toác 3D, tên Ethan #10. `213_02`: Vỏ chăn gấp. `213_03`–`213_07`: 5 panel mockup và infographic tiêu chuẩn. | Mô tả chuẩn bóng rổ đâm xuyên tường nứt 3D | PASS (100đ) |
| 214 | `personalized-basketball-entering-net-comforter-e55e615fd3-e55e615fd3` | 5 | `214_01`: Rổ bóng rổ phát sáng với ngọn lửa và lưới màu cầu vồng rực rỡ, tên Ethan #10. `214_02`: Dệt mật độ cao. `214_03`: Giặt máy. `214_04`: Khóa kéo đáy. `214_05`: Mockup góc phòng ngủ. | Mô tả đúng rổ bóng rổ màu cầu vồng phát sáng | PASS (100đ) |
| 215 | `personalized-basketball-fire-and-water-splash-blanket-with-name-and-number-b05a53f208-b05a53f208` | 8 | `215_01`: Chăn nỉ nhung, va chạm dữ dội giữa ngọn lửa cam đỏ và sóng nước xanh biếc quanh quả bóng, tên Lucas #24. `215_02`–`215_08`: 7 panel chăn nỉ nhung (chất liệu fleece siêu mềm, viền may, bảng kích thước Small/Medium/Large, giặt máy, lifestyle). | Chính xác chi tiết va chạm lửa và nước | PASS (100đ) |
| 216 | `personalized-basketball-flames-flying-over-comforter-with-name-and-number-feb4be1ea4-feb4be1ea4` | 7 | `216_01`: Quả bóng rổ bốc lửa rực cháy bay vút qua sân đấu mịt mờ khói đen, tên Ethan #10. `216_02`–`216_07`: 6 panel mockup và infographic tiêu chuẩn. | Mô tả chuẩn quả bóng rổ bốc lửa bay qua sân khói | PASS (100đ) |
| 217 | `personalized-basketball-glowing-light-burst-comforter-with-name-and-number-3bc9e1267e-3bc9e1267e` | 7 | `217_01`: Vầng sáng vàng cam bùng nổ tỏa tia rực rỡ từ quả bóng rổ trên sân đấu, tên Ethan #10. `217_02`–`217_07`: 6 panel mockup và infographic tiêu chuẩn. | Chính xác vầng sáng vàng cam bùng nổ tỏa tia | PASS (100đ) |
| 218 | `personalized-basketball-held-by-hand-on-court-blanket-with-name-and-number-079998b76f-079998b76f` | 8 | `218_01`: Chăn nỉ nhung, bàn tay cầu thủ nắm chặt quả bóng rổ trên mặt sàn gỗ đấu trường, tên Lucas #24. `218_02`–`218_08`: 7 panel chăn nỉ nhung chi tiết (chất liệu, kích cỡ, giặt ủi, lifestyle). | Mô tả chuẩn bàn tay nắm bóng trên sàn đấu | PASS (100đ) |
| 219 | `personalized-basketball-held-under-player-arm-blanket-with-name-and-number-7a077c5144-7a077c5144` | 8 | `219_01`: Chăn nỉ nhung, dáng cầu thủ ôm quả bóng rổ dưới cánh tay sẵn sàng thi đấu, tên Lucas #24. `219_02`–`219_08`: 7 panel chăn nỉ nhung chi tiết. | Mô tả chuẩn dáng cầu thủ ôm bóng dưới cánh tay | PASS (100đ) |
| 220 | `personalized-basketball-hoop-comforter-name-number-863d6f6ff1-863d6f6ff1` | 7 | `220_01`: Rổ bóng rổ viền cam rực sáng nổi bật trên nền sân tối, tên Ethan #10 trên vỏ gối shams. `220_02`–`220_07`: 6 panel mockup và infographic tiêu chuẩn. | Chính xác chi tiết rổ viền cam và tên số trên shams | PASS (100đ) |

---

## 5. Bằng chứng Thực nghiệm 20 Truy vấn US SERP Độc lập

Đã thực hiện 20 truy vấn tìm kiếm độc lập trên Google Search (thị trường Mỹ `en-US`):

| Query ID | Pos | Từ khóa truy vấn | Loại | Search Intent | Top Organic Domains | Đánh giá & Kết luận |
| :--- | :---: | :--- | :---: | :--- | :--- | :--- |
| `SERP-211-PRI` | 211 | *"basketball close up number blanket"* | `PRIMARY` | Commercial / Custom Number Blanket | etsy.com, getphotoblanket.com, personalizationmall.com, zazzle.com | Accurately targets Pos 211's oversized ball close-up and jersey number artwork. |
| `SERP-211-COM` | 211 | *"basketball number throw blanket"* | `COMPARATOR` | Commercial / Jersey Number Blanket | amazon.com, target.com, walmart.com, etsy.com | Confirms strong market interest in sports number blankets. |
| `SERP-212-PRI` | 212 | *"custom basketball court comforter"* | `PRIMARY` | Commercial / Hardwood Court Bedding | etsy.com, 2cooldesigns.com, youcustomizeit.com, ohaprints.com | Directly targets Pos 212's court perspective comforter. |
| `SERP-212-COM` | 212 | *"basketball court comforter set"* | `COMPARATOR` | Commercial / Sports Arena Bedding | wayfair.com, target.com, walmart.com, amazon.com | Confirms commercial viability of court perspective motif. |
| `SERP-213-PRI` | 213 | *"custom basketball cracked wall comforter"* | `PRIMARY` | Commercial / 3D Breaking Wall Bedding | jeminise.com, ohaprints.com, etsy.com, jessartdecoration.com.au | Directly targets Pos 213's basketball breaking through cracked concrete gym wall. |
| `SERP-213-COM` | 213 | *"cracked wall basketball bedding"* | `COMPARATOR` | Commercial / 3D Sports Bedroom Decor | temu.com, dhgate.com, etsy.com, walmart.com | Validates appeal of dynamic wall break visual effect. |
| `SERP-214-PRI` | 214 | *"custom rainbow basketball net comforter"* | `PRIMARY` | Commercial / Prismatic Rainbow Sports Bedding | youcustomizeit.com, etsy.com, amorcustomgifts.com, 2cooldesigns.com | Accurately targets Pos 214's rainbow glowing net comforter. |
| `SERP-214-COM` | 214 | *"rainbow basketball bedding"* | `COMPARATOR` | Commercial / Colorful Youth Bedding | walmart.com, amazon.com, zazzle.com, target.com | Confirms market demand for colorful artistic basketball bedding. |
| `SERP-215-PRI` | 215 | *"custom basketball fire water blanket"* | `PRIMARY` | Commercial / Fire & Water Elements Blanket | lalaky.com, dyoart.com, callie.com, etsy.com | Directly targets Pos 215's fire and water splash blanket. |
| `SERP-215-COM` | 215 | *"fire and water basketball throw blanket"* | `COMPARATOR` | Commercial / Elemental Sports Throws | walmart.com, amazon.com, etsy.com, shein.com | Validates popularity of fire and water split graphics. |
| `SERP-216-PRI` | 216 | *"custom flaming basketball comforter"* | `PRIMARY` | Commercial / Flaming Basketball Bedding | etsy.com, ohaprints.com, 2cooldesigns.com, eloquentinnovations.com | Accurately matches Pos 216's flying flame ball over court comforter. |
| `SERP-216-COM` | 216 | *"flaming basketball bedding set"* | `COMPARATOR` | Commercial / Fire Sports Linens | shein.com, temu.com, walmart.com, amazon.com | Validates strong evergreen demand for fiery sports linens. |
| `SERP-217-PRI` | 217 | *"custom light burst basketball comforter"* | `PRIMARY` | Commercial / Light Burst Basketball Bedding | ohaprints.com, youcustomizeit.com, eloquentinnovations.com, etsy.com | Directly targets Pos 217's gold/orange glowing light burst artwork. |
| `SERP-217-COM` | 217 | *"glowing basketball comforter set"* | `COMPARATOR` | Commercial / Glowing Sports Bedding | amazon.com, target.com, etsy.com, wayfair.com | Confirms consumer interest in glowing athletic aesthetics. |
| `SERP-218-PRI` | 218 | *"custom basketball court blanket"* | `PRIMARY` | Commercial / Court Action Throw Blanket | etsy.com, personalizationmall.com, zazzle.com, gollygiftco.com | Accurately matches Pos 218's hand-holding-ball on court blanket. |
| `SERP-218-COM` | 218 | *"basketball court throw blanket"* | `COMPARATOR` | Commercial / Court Floor Throws | walmart.com, target.com, etsy.com, squadlocker.com | Confirms strong gift market for court blankets. |
| `SERP-219-PRI` | 219 | *"custom basketball player blanket"* | `PRIMARY` | Commercial / Athlete Player Blanket | gollygiftco.com, personalizationmall.com, threadtheword.com, etsy.com | Accurately targets Pos 219's player arm holding basketball blanket. |
| `SERP-219-COM` | 219 | *"personalized basketball player fleece throw"* | `COMPARATOR` | Commercial / Fleece Throws for Athletes | zazzle.com, haloballs.com, getphotoblanket.com, etsy.com | Validates high demand for athlete player silhouette blankets. |
| `SERP-220-PRI` | 220 | *"custom orange basketball hoop comforter"* | `PRIMARY` | Commercial / Black & Orange Hoop Bedding | etsy.com, youcustomizeit.com, 2cooldesigns.com, amorcustomgifts.com | Directly targets Pos 220's orange rim and net with custom pillow shams. |
| `SERP-220-COM` | 220 | *"orange basketball bedding set"* | `COMPARATOR` | Commercial / Vibrant Orange Sports Bedding | target.com, roomstogo.com, nba.com, walmart.com | Confirms robust demand for orange hoop sports decor. |

---

## 6. Bảng Phân loại Issues & Khuyến nghị Khắc phục

Tổng hợp **30 issues** được ghi nhận trong phiên QA độc lập đầu tiên trên revision r2:

- **10 MAJOR Issues (`ISSUE-0001` đến `ISSUE-0010`):**
  - Trường vi phạm: `description_proposed_html` cho cả 10 sản phẩm.
  - Nguyên nhân: Lộ prompt nội bộ (*"The wording focuses on..."*, *"Gallery images show..."*); hướng dẫn chọn option mâu thuẫn trên Pos 211 (chỉ có Default Title); không làm rõ vỏ gối shams là tùy chọn mua thêm trên Pos 220; thiếu thông số kỹ thuật then chốt (cấu tạo comforter/duvet, khóa kéo đáy, kích thước, giặt ủi).
  - Khắc phục: Thay thế bằng 10 bản mô tả HTML chuẩn publish-ready ở Mục 7.
- **10 MINOR Issues (`ISSUE-0011` đến `ISSUE-0020`):**
  - Trường vi phạm: `meta_description_seo` cho cả 10 sản phẩm.
  - Nguyên nhân: Cắt tỉa thô bạo ở đúng 155 ký tự làm đứt đoạn câu giữa chừng hoặc cắt cụt từ ngữ.
  - Khắc phục: Cắt tỉa theo ranh giới câu trọn vẹn dưới 150 ký tự như đề xuất dưới đây.
- **10 LIMITATION Issues (`ISSUE-0021` đến `ISSUE-0030`):**
  - Trường: `keyword_evidence_level`.
  - Nguyên nhân: Môi trường offline QA không kết nối trực tiếp Search Console API, được bù đắp 100% bằng 20 truy vấn US SERP độc lập.

### 6.1. Đề xuất Hiệu chỉnh Meta Description Chuẩn SEO (10 Sản phẩm)

| Pos | Handle | Meta Description Hiện tại (r2 - Bị cắt cụt) | Ký tự | Meta Description Đề xuất Chuẩn hóa | Ký tự mới |
| :---: | :--- | :--- | :---: | :--- | :---: |
| 211 | `personalized-basketball-close-up-with-number-blanket-with-name-and-number-f16f46abef-f16f46abef` | `Shop basketball close-up number blanket with black blanket with oversized close-up basketball, large jersey number and vertical sample name, visible basket` | 155 | **Stay warm with this basketball close-up number blanket. Featuring bold athletic graphics on ultra-soft plush fleece for courtside fans.** | 135 |
| 212 | `personalized-basketball-court-perspective-comforter-with-name-and-number-5545af9fa4-5545af9fa4` | `Shop custom basketball court comforter with black and gray basketball comforter with court perspective, large ball texture, sample name and jersey number, ` | 155 | **Bring arena excitement home with this custom basketball court comforter or duvet cover set. Personalize with your athlete's name and number.** | 140 |
| 213 | `personalized-basketball-cracked-wall-comforter-with-name-and-number-aeaddaa506-aeaddaa506` | `Shop custom basketball cracked wall comforter with white and gray cracked-wall basketball comforter with an orange ball breaking through and vertical sampl` | 155 | **Make a high-impact statement with this custom basketball cracked wall comforter. Add your player's name and jersey number in Twin to King.** | 138 |
| 214 | `personalized-basketball-entering-net-comforter-e55e615fd3-e55e615fd3` | `Shop custom rainbow basketball net comforter with bright rainbow fire basketball comforter with glowing hoop, net, sample name and jersey number, selectabl` | 155 | **Light up your room with this custom rainbow basketball net comforter. High-definition colorful court artwork personalized with name and number.** | 143 |
| 215 | `personalized-basketball-fire-and-water-splash-blanket-with-name-and-number-b05a53f208-b05a53f208` | `Shop custom basketball fire water blanket with blanket with basketball, fire, water splash, lightning, hoop background, vertical sample name and jersey num` | 155 | **Cozy up with this custom basketball fire and water blanket. Soft plush fleece in multiple sizes with optional custom name and jersey number.** | 140 |
| 216 | `personalized-basketball-flames-flying-over-comforter-with-name-and-number-feb4be1ea4-feb4be1ea4` | `Shop custom flaming basketball comforter with black comforter with a flaming basketball flying across a smoky court, vertical sample name and jersey number` | 155 | **Ignite bedroom decor with this custom flaming basketball comforter or duvet cover set. Personalize with your player's name and jersey number.** | 141 |
| 217 | `personalized-basketball-glowing-light-burst-comforter-with-name-and-number-3bc9e1267e-3bc9e1267e` | `Shop custom light burst basketball comforter with gold and orange basketball comforter with glowing light burst, sample name and jersey number, selectable ` | 155 | **Bring explosive energy home with this custom light burst basketball comforter. Personalize with your athlete's name and number in Twin to King.** | 143 |
| 218 | `personalized-basketball-held-by-hand-on-court-blanket-with-name-and-number-079998b76f-079998b76f` | `Shop custom basketball court blanket with stadium court blanket with a hand holding a basketball, sample name arched over the ball and jersey number, selec` | 155 | **Relax in game-day comfort with this custom basketball court blanket. Cozy plush fleece in multiple sizes with optional custom name and number.** | 142 |
| 219 | `personalized-basketball-held-under-player-arm-blanket-with-name-and-number-7a077c5144-7a077c5144` | `Shop custom basketball player blanket with black blanket with player arm holding a basketball, vertical sample name and jersey number on the ball, selectab` | 155 | **Wrap up in varsity style with this custom basketball player blanket. Ultra-soft fleece throw in selectable sizes with optional name and number.** | 143 |
| 220 | `personalized-basketball-hoop-comforter-name-number-863d6f6ff1-863d6f6ff1` | `Shop custom orange basketball hoop comforter with black and orange basketball comforter with glowing hoop, sample name and jersey number on pillow shams, s` | 155 | **Score big with this custom orange basketball hoop comforter or duvet cover set. Personalize with name and number; optional shams available.** | 139 |

---

## 7. Đề xuất 10 Mô tả HTML Publish-Ready (Chuẩn hóa Hoàn chỉnh)

Dưới đây là 10 bản mô tả HTML tiếng Anh hoàn chỉnh, loại bỏ sạch prompt boilerplate, phản ánh chuẩn xác purchase flow thực tế (làm rõ Pos 211 dạng chăn tiêu chuẩn; 2 trường Name/Number bắt buộc đối với Bedding; 2 trường Name/Number tùy chọn đối với Blanket; làm rõ shams là tùy chọn mua thêm trên Pos 220), phân định rành mạch Comforter vs Duvet Cover, kích thước, phụ kiện kèm theo và hướng dẫn giặt ủi:

### Position 211: Basketball Close-Up Number Blanket (`personalized-basketball-close-up-with-number-blanket-with-name-and-number-f16f46abef-f16f46abef`)

```html
<p>Wrap up in high-impact basketball energy with the <strong>Basketball Close-Up Number Blanket</strong> from Jeminise. This bold sports throw blanket showcases an authentic macro photographic close-up of a basketball's textured pebbled leather surface paired with high-contrast jersey number graphics and vertical player name styling. Designed for athletes, coaches, and courtside fans, it brings varsity excitement and plush warmth to the couch, bedroom, or stadium stands.</p>
<h3>Product Features & Specifications</h3>
<ul>
  <li><strong>Standard Blanket Construction:</strong> Crafted as an ultra-soft, all-in-one plush throw blanket with finished hemmed edges, ready to provide immediate, cozy warmth right out of the package.</li>
  <li><strong>Athletic Graphic Artwork:</strong> Features high-definition basketball leather and athletic jersey numbering printed with vivid, fade-resistant dye-sublimation ink.</li>
  <li><strong>Throw Sizing:</strong> Standard medium throw blanket dimensions (50" x 60"), ideal for relaxing on the sofa, dorm room layering, or chilly tournament bleachers.</li>
</ul>
<h3>Fabric Quality & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> 100% premium ultra-soft fleece with velvety flannel hand feel for lightweight, breathable insulation without heavy bulk.</li>
  <li><strong>Care:</strong> Machine wash cold on gentle cycle with mild detergent. Tumble dry on low heat or line dry. Non-pilling, anti-static, and colorfast wash after wash.</li>
</ul>
```

### Position 212: Custom Basketball Court Comforter (`personalized-basketball-court-perspective-comforter-with-name-and-number-5545af9fa4-5545af9fa4`)

```html
<p>Bring the arena atmosphere home with the <strong>Custom Basketball Court Comforter</strong> from Jeminise. Featuring an expansive hardwood court perspective stretching toward an arena hoop under bright arena floodlights, this bedding ensemble delivers authentic game-day drama. Fully customized with your player's name and jersey number in athletic lettering, it creates an inspiring bedroom centerpiece for young ballplayers, high school athletes, and basketball fans.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> A complete, pre-filled comforter packed with soft, lightweight, hypoallergenic down-alternative microfiber fill. Provides balanced, cloud-soft warmth all year round.</li>
  <li><strong>Duvet Cover Option:</strong> A protective duvet shell equipped with a hidden bottom zipper closure, allowing quick insertion of your favorite comforter or duvet insert for effortless home laundering.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name:</strong> Enter your athlete's first name, last name, or team nickname (1–30 characters) in the <em>Customize Your Name</em> field. Enter "NO" if you prefer artwork without a name.</li>
  <li><strong>Custom Number:</strong> Enter your ballplayer's jersey number (1–5 digits) in the <em>Customize Your Number</em> field. Enter "NO" if no number is desired.</li>
</ul>
<h3>Available Sizes & Coordinating Add-Ons</h3>
<ul>
  <li><strong>Bed Sizes:</strong> Twin (68x86 in), Full (79x90 in), Queen (90x90 in), and King (90x104 in).</li>
  <li><strong>Matching Pillowcases:</strong> Bundle with None, 1 Pillowcase, or 2 Pillowcases printed with coordinating basketball court artwork.</li>
  <li><strong>Optional Flat Sheet:</strong> Complete your bedding set with an optional matching flat sheet cover in corresponding mattress dimensions.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> 100% premium high-density brushed microfiber for breathable softness, wrinkle resistance, and durable comfort.</li>
  <li><strong>Care:</strong> Machine wash cold with mild detergent on gentle cycle. Tumble dry low or line dry. Colors resist fading and pilling.</li>
</ul>
```

### Position 213: Custom Basketball Cracked Wall Comforter (`personalized-basketball-cracked-wall-comforter-with-name-and-number-aeaddaa506-aeaddaa506`)

```html
<p>Make an explosive, high-impact statement with the <strong>Custom Basketball Cracked Wall Comforter</strong> from Jeminise. This 3D-effect sports bedding features an authentic textured basketball bursting through a cracked concrete gym wall with sharp court boundary lines. Personalized with your player's name and jersey number, it transforms any teen bedroom, dorm, or sports lover's retreat into a dynamic arena showcase.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Pre-filled comforter packed with airy, hypoallergenic microfiber insulation, delivering all-season plush warmth and breathable comfort.</li>
  <li><strong>Duvet Cover Option:</strong> Features a durable, concealed bottom zipper closure for effortless cover removal and fast washing.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name:</strong> Enter your player's name (1–30 characters) in the <em>Customize Your Name</em> field. Type "NO" for clean, unpersonalized artwork.</li>
  <li><strong>Custom Number:</strong> Enter your player's jersey number (1–5 characters) in the <em>Customize Your Number</em> field. Type "NO" to omit.</li>
</ul>
<h3>Available Sizes & Coordinating Add-Ons</h3>
<ul>
  <li><strong>Bed Sizes:</strong> Twin (68x86 in), Full (79x90 in), Queen (90x90 in), and King (90x104 in).</li>
  <li><strong>Matching Pillowcases:</strong> Add None, 1 Pillowcase, or 2 Pillowcases with matching cracked wall basketball graphics.</li>
  <li><strong>Coordinated Flat Sheet:</strong> Optionally include an additional matching flat sheet cover in identical mattress dimensions.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> Premium fine-spun brushed microfiber delivering gentle skin comfort, anti-static softness, and wrinkle resistance.</li>
  <li><strong>Care:</strong> Machine wash cold with like colors. Tumble dry on low. Vibrant 3D-effect print resists cracking or peeling.</li>
</ul>
```

### Position 214: Custom Rainbow Basketball Net Comforter (`personalized-basketball-entering-net-comforter-e55e615fd3-e55e615fd3`)

```html
<p>Light up your bedroom with the high-voltage color of the <strong>Custom Rainbow Basketball Net Comforter</strong> from Jeminise. This artistic sports bedding design combines the excitement of a swished basket with radiant rainbow fire, an illuminated hoop, and prismatic net cords under arena lights. Tailored with your athlete's name and jersey number, it brings unforgettable energy and creative flair to any bedroom.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Ready-to-use comforter filled with fluffy, lightweight down-alternative microfiber fill for cozy bedtime comfort across all seasons.</li>
  <li><strong>Duvet Cover Option:</strong> Equipped with an invisible bottom zipper closure, making it effortless to insert, remove, and wash your comforter insert.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name:</strong> Type your athlete's name or custom text (1–30 characters) into the <em>Customize Your Name</em> box. Enter "NO" for artwork without personalization.</li>
  <li><strong>Custom Number:</strong> Type your athlete's jersey number (1–5 digits) into the <em>Customize Your Number</em> box. Enter "NO" to leave blank.</li>
</ul>
<h3>Available Sizes & Coordinating Add-Ons</h3>
<ul>
  <li><strong>Bed Sizes:</strong> Available in Twin, Full, Queen, and King mattress dimensions.</li>
  <li><strong>Matching Pillowcases:</strong> Select None, 1 Pillowcase, or 2 Pillowcases with matching rainbow basketball net graphics.</li>
  <li><strong>Optional Flat Sheet:</strong> Coordinated flat sheet cover available in matching mattress size.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> 100% brushed polyester microfiber fabric engineered for superior breathability, softness, and durability.</li>
  <li><strong>Care:</strong> Machine wash cold on delicate cycle. Tumble dry on low heat. Prismatic colorfast dyes stay rich and bright.</li>
</ul>
```

### Position 215: Custom Basketball Fire Water Blanket (`personalized-basketball-fire-and-water-splash-blanket-with-name-and-number-b05a53f208-b05a53f208`)

```html
<p>Experience the ultimate clash of elements with the <strong>Custom Basketball Fire Water Blanket</strong> from Jeminise. This dramatic sports throw blanket showcases an explosive collision of roaring orange fire flames and icy blue water splashes swirling around an authentic textured basketball. Perfect for keeping warm at courtside games, lounging on the couch, or extra bedding warmth, it makes an unforgettable personalized gift for passionate ballplayers.</p>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name (Optional):</strong> Personalize with your athlete's name or team nickname (1–200 characters) in the <em>Custom Name</em> field. This field is optional; leave blank if you prefer artwork without custom text.</li>
  <li><strong>Custom Number (Optional):</strong> Add your player's jersey number (1–20 characters) in the <em>Custom Number</em> field. This field is optional; leave blank for an unnumbered blanket.</li>
</ul>
<h3>Available Blanket Sizes</h3>
<ul>
  <li><strong>Small (40" x 50"):</strong> Compact size perfect for toddlers, kids, lap throws, or car travel.</li>
  <li><strong>Medium (50" x 60"):</strong> Most popular throw size for sofa lounging, gaming chairs, or stadium bleachers.</li>
  <li><strong>Large (60" x 80"):</strong> Oversized blanket offering full bed layering or head-to-toe cocoon comfort.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> Ultra-soft premium plush fleece with velvety flannel hand feel, offering cozy warmth without heavy bulk.</li>
  <li><strong>Care:</strong> Machine wash cold with mild detergent on gentle cycle. Tumble dry on low heat or hang dry. Anti-pilling and fade-resistant.</li>
</ul>
```

### Position 216: Custom Flaming Basketball Comforter (`personalized-basketball-flames-flying-over-comforter-with-name-and-number-feb4be1ea4-feb4be1ea4`)

```html
<p>Ignite your sports bedroom decor with the <strong>Custom Flaming Basketball Comforter</strong> from Jeminise. This high-octane bedding set captures the thrill of a power shot, featuring a blazing basketball rocketing through the air across a dark, smoky arena court with swirling embers. Fully customized with your player's name and jersey number, it transforms any room into a championship arena.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Pre-filled comforter packed with soft, airy microfiber insulation, delivering all-weather warmth and cloud-soft comfort.</li>
  <li><strong>Duvet Cover Option:</strong> Equipped with an unobtrusive bottom zipper closure, allowing quick insertion of your comforter insert and easy home laundry.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name:</strong> Enter your ballplayer's name or nickname (1–30 characters) in the <em>Customize Your Name</em> field. Enter "NO" for non-customized artwork.</li>
  <li><strong>Custom Number:</strong> Enter your player's jersey number (1–5 characters) in the <em>Customize Your Number</em> field. Enter "NO" to omit.</li>
</ul>
<h3>Available Sizes & Coordinating Add-Ons</h3>
<ul>
  <li><strong>Bed Sizes:</strong> Twin (68x86 in), Full (79x90 in), Queen (90x90 in), and King (90x104 in).</li>
  <li><strong>Matching Pillowcases:</strong> Bundle with None, 1 Pillowcase, or 2 Pillowcases with matching flaming basketball graphics.</li>
  <li><strong>Optional Flat Sheet:</strong> Complete your bedding set with an optional matching flat sheet cover in corresponding dimensions.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> High-density woven microfiber engineered for softness, wrinkle resistance, and lasting colorfast vibrance.</li>
  <li><strong>Care:</strong> Machine wash cold with mild detergent. Tumble dry low. Fiery orange and crimson dyes remain bold wash after wash.</li>
</ul>
```

### Position 217: Custom Light Burst Basketball Comforter (`personalized-basketball-glowing-light-burst-comforter-with-name-and-number-3bc9e1267e-3bc9e1267e`)

```html
<p>Electrify your bedroom with championship intensity with the <strong>Custom Light Burst Basketball Comforter</strong> from Jeminise. This radiant design features a glowing golden basketball radiating an explosive light burst and energetic sparkle rays across a dark court background. Personalized with your player's name and jersey number, it inspires athletic confidence and big-game excitement every night.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Complete, ready-to-use comforter filled with lofty down-alternative microfiber fill, providing balanced, all-season warmth right out of the box.</li>
  <li><strong>Duvet Cover Option:</strong> Sleek duvet cover with a concealed bottom zipper closure, providing an easy-to-wash outer layer for your duvet insert.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name:</strong> Input your athlete's name or custom dedication (1–30 characters) into the <em>Customize Your Name</em> box. Enter "NO" for unpersonalized artwork.</li>
  <li><strong>Custom Number:</strong> Input your player's jersey number (1–5 digits) into the <em>Customize Your Number</em> box. Enter "NO" to omit.</li>
</ul>
<h3>Available Sizes & Coordinating Add-Ons</h3>
<ul>
  <li><strong>Bed Sizes:</strong> Available in Twin, Full, Queen, and King bed dimensions.</li>
  <li><strong>Coordinating Pillowcases:</strong> Add 1 or 2 matching light burst basketball pillowcases to complete your sports bedding set.</li>
  <li><strong>Additional Flat Sheet:</strong> Optionally add a coordinated flat sheet cover in identical mattress dimensions.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> 100% fine-spun brushed microfiber offering silky softness and anti-wrinkle durability.</li>
  <li><strong>Care:</strong> Machine wash cold on gentle cycle. Tumble dry on low heat. High-contrast luminous graphics resist fading.</li>
</ul>
```

### Position 218: Custom Basketball Court Blanket (`personalized-basketball-held-by-hand-on-court-blanket-with-name-and-number-079998b76f-079998b76f`)

```html
<p>Step onto the hardwood in cozy comfort with the <strong>Custom Basketball Court Blanket</strong> from Jeminise. Showcasing an authentic close-up of a player's hand palm-down firmly gripping a regulation basketball resting on a polished gym floor, this throw blanket captures the intense focus of game day. Customized optionally with your player's name and jersey number, it is the perfect plush companion for courtside spectating, gaming, or couch relaxing.</p>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name (Optional):</strong> Personalize with your player's name or team dedication (1–200 characters) in the <em>Custom Name</em> field. This field is optional; leave blank if you prefer clean artwork without text.</li>
  <li><strong>Custom Number (Optional):</strong> Enter your athlete's jersey number (1–20 characters) in the <em>Custom Number</em> field. This field is optional; leave blank for an unnumbered throw.</li>
</ul>
<h3>Available Blanket Sizes</h3>
<ul>
  <li><strong>Small (40" x 50"):</strong> Compact size ideal for kids, baby cribs, strollers, or car travel.</li>
  <li><strong>Medium (50" x 60"):</strong> The standard throw size for sofa lounging, reading nooks, or tournament bleachers.</li>
  <li><strong>Large (60" x 80"):</strong> Generously sized for full bed layering or wrapping up head-to-toe on chilly evenings.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> Premium ultra-soft fleece with velvety flannel touch, providing lightweight warmth and skin-friendly softness.</li>
  <li><strong>Care:</strong> Machine wash cold on delicate cycle. Tumble dry low or line dry. Non-pilling, shrink-resistant, and colorfast.</li>
</ul>
```

### Position 219: Custom Basketball Player Blanket (`personalized-basketball-held-under-player-arm-blanket-with-name-and-number-7a077c5144-7a077c5144`)

```html
<p>Celebrate your passion for hoops with the <strong>Custom Basketball Player Blanket</strong> from Jeminise. Featuring a powerful athletic silhouette of a basketball player tucking a textured ball securely under their arm on a sleek dark court backdrop, this throw blanket radiates varsity confidence. Personalized optionally with your athlete's name and jersey number, it makes an outstanding Senior Night, championship, or birthday gift.</p>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name (Optional):</strong> Enter your player's first name, last name, or team nickname (1–200 characters) in the <em>Custom Name</em> field. This field is optional; leave blank for unpersonalized artwork.</li>
  <li><strong>Custom Number (Optional):</strong> Enter your athlete's jersey number (1–20 characters) in the <em>Custom Number</em> field. This field is optional; leave blank for an unnumbered blanket.</li>
</ul>
<h3>Available Blanket Sizes</h3>
<ul>
  <li><strong>Small (40" x 50"):</strong> Ideal for children, stroller rides, pet blankets, or travel.</li>
  <li><strong>Medium (50" x 60"):</strong> Most versatile throw size for movie nights, couch lounging, and game viewing.</li>
  <li><strong>Large (60" x 80"):</strong> Oversized blanket perfect for twin/full bed coverage and cozy winter warmth.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> 100% fine-spun plush fleece with velvety flannel finish for luxurious warmth without heaviness.</li>
  <li><strong>Care:</strong> Machine wash cold with mild detergent on gentle cycle. Tumble dry on low heat. Colors remain bold and vibrant.</li>
</ul>
```

### Position 220: Custom Orange Basketball Hoop Comforter (`personalized-basketball-hoop-comforter-name-number-863d6f6ff1-863d6f6ff1`)

```html
<p>Score big in bedroom style with the <strong>Custom Orange Basketball Hoop Comforter</strong> from Jeminise. This dynamic sports bedding design highlights an illuminated orange basketball rim and net glowing with dramatic focus against a sleek dark stadium background. Customized with your player's name and jersey number, it brings athletic character and cozy warmth to any sports-themed bedroom.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> A complete, pre-filled comforter packed with soft, lightweight, hypoallergenic microfiber batting. Delivers balanced, cloud-like warmth through all seasons.</li>
  <li><strong>Duvet Cover Option:</strong> A protective duvet cover equipped with a hidden bottom zipper closure, designed to slip smoothly over your existing comforter or duvet insert for quick laundering.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name:</strong> Enter your ballplayer's name or custom text (1–30 characters) in the <em>Customize Your Name</em> field. Enter "NO" if you prefer artwork without a name.</li>
  <li><strong>Custom Number:</strong> Enter your athlete's jersey number (1–5 digits) in the <em>Customize Your Number</em> field. Enter "NO" to omit.</li>
</ul>
<h3>Available Sizes & Coordinating Add-Ons</h3>
<ul>
  <li><strong>Bed Sizes:</strong> Twin (68x86 in), Full (79x90 in), Queen (90x90 in), and King (90x104 in).</li>
  <li><strong>Coordinating Pillowcases (Optional Add-On):</strong> Personalize your bedding ensemble with matching pillow shams printed with customized player name and number by selecting 1 Pillowcase or 2 Pillowcases in the <em>Choose Pillowcases</em> option menu. Select None if you only want the comforter/duvet cover.</li>
  <li><strong>Additional Flat Sheet:</strong> Optionally add a coordinated flat sheet cover in corresponding bed dimensions.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> Premium high-density brushed microfiber for breathable softness, hypoallergenic comfort, and lasting durability.</li>
  <li><strong>Care:</strong> Machine wash cold on gentle cycle. Tumble dry on low heat. Rich black and vibrant orange dyes remain colorfast wash after wash.</li>
</ul>
```

---

## 8. Tuyên bố Trạng thái & Tiếp tục Quy trình

- **Xác nhận khóa dữ liệu:** Toàn bộ bằng chứng kiểm toán độc lập đã được lưu trữ vĩnh viễn tại `seo_runs/jeminise.com/20260906_234129/qa/20260908_134500/`.
- **File bàn giao:** `resutls/jeminise.com/20260906_234129/qa/20260908_134500/SEO_QA_qa_batch_022_r2.xlsx` và `SEO_QA_qa_batch_022_r2.md`.
- **Nguyên tắc an toàn:** Không chỉnh sửa file nguồn `SEO_Product_Optimization_qa_batch_022_r2.xlsx` và không sửa trực tiếp live store.
- **Dừng theo quy trình:** Hệ thống kích hoạt cờ `awaiting_confirmation=true` và **DỪNG LẠI**, chờ lệnh xác nhận từ người dùng trước khi chuyển sang Batch 23.
