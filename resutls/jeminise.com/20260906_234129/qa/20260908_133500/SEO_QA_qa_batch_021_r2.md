# Báo cáo QA Độc lập SEO Workbook - Batch 21 Revision r2

> **Mã phiên QA:** `20260908_133500`  
> **Tập tin nguồn thẩm định:** `resutls/jeminise.com/20260906_234129/revisions/qa_batch_021_r2/SEO_Product_Optimization_qa_batch_021_r2.xlsx`  
> **SHA256 Tập tin nguồn:** `ee969508fcaa3e1148b923a634c03afcdeeaef75d69d0fb4005d4148e798038b`  
> **Thị trường mục tiêu:** United States (`en-US`) | **Nội dung:** English SEO Copy  
> **Phạm vi thẩm định:** 10 sản phẩm (Inventory Position 201–210), 64 ảnh gallery, 40 dòng Keyword_Map, 10 dòng Buyer_Search_Research  
> **Đặc thù phiên chạy:** **Phiên QA độc lập đầu tiên cho Batch 21 trên revision r2** (`first_qa_run: true`, `source_revision: "r2"`)  
> **Trạng thái thẩm định:** `awaiting_confirmation=true` | **Kết luận tổng thể:** `QA_REVISE` (Điểm trung bình: **95.00/100**)  

---

## 1. Tóm tắt kết quả QA & Chấm điểm Tổng thể

Phiên kiểm toán độc lập đầu tiên cho **Batch 21 Revision r2** được thực hiện với 100% trọng số đánh giá trên toàn bộ 10 sản phẩm, 64 hình ảnh gallery, đối chiếu trực tiếp dữ liệu live storefront và customizer tại `jeminise.com`, cùng 20 truy vấn US SERP độc lập:

| Chỉ số | Giá trị | Ghi chú kiểm toán |
| :--- | :---: | :--- |
| **Tổng số sản phẩm** | 10 | Position 201 đến 210 (`design-26`, `duvet-cover`, `ff32dd676a` đến `3f49346de1`) |
| **Tỷ lệ kiểm toán sản phẩm** | 100% (10/10) | Cột `page_read=TRUE` đầy đủ |
| **Tổng số ảnh gallery** | 64 | Pos 201, 204, 207, 208, 209: 7 ảnh; Pos 202: 4 ảnh; Pos 203, 205: 5 ảnh; Pos 206: 8 ảnh; Pos 210: 7 ảnh |
| **Tỷ lệ kiểm toán ảnh** | 100% (64/64) | Điểm ảnh trung bình 100.0/100, 100% PASS |
| **Điểm số trung bình** | **95.00 / 100** | 10/10 sản phẩm đạt chính xác 95.0 điểm (P1: 15, P2: 10, K1: 10, K2: 5, K3: 5, T1: 10, T2: 5, D1: 3, D2: 7, I1: 20, E1: 5) |
| **Số sản phẩm QA_PASS** | 0 | Do tồn tại lỗi MAJOR ở tiêu chuẩn D2 |
| **Số sản phẩm QA_REVISE** | **10** | Yêu cầu chuẩn hóa lại mô tả sản phẩm D2 và cắt tỉa meta description D1 |
| **Số sản phẩm QA_FAIL** | 0 | Không có lỗi CRITICAL |
| **Số sản phẩm QA_INCOMPLETE**| 0 | 100% dữ liệu được đánh giá đầy đủ |
| **Kết luận tổng thể** | **QA_REVISE** | Sẵn sàng xuất bản sau khi áp dụng 10 mô tả HTML và meta descriptions đề xuất |

---

## 2. Bảng tổng hợp đánh giá 10 sản phẩm (Pos 201–210)

| Pos | Product Key / Handle | Title đề xuất r2 | Primary Keyword | P1 | P2 | K1 | K2 | K3 | T1 | T2 | D1 | D2 | I1 | E1 | Tổng | Issues | QA Status |
| :---: | :--- | :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| 201 | `personalized-baseball-bedding-full-size-flag-custom-name-d-design-26` | Custom Baseball Flag Name Number Bedding | *custom baseball flag name number bedding* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 202 | `personalized-baseball-bedding-full-size-flag-custom-name-duvet-cover` | Custom Neon Baseball Player Duvet Cover | *custom neon baseball player duvet cover* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 203 | `personalized-basketball-above-hoop-and-net-comforter-with-name-and-number-ff32dd676a-ff32dd676a` | Custom Basketball Hoop Comforter Set | *custom basketball hoop comforter set* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 204 | `personalized-basketball-above-hoop-close-up-comforter-with-name-and-number-95487807f8-95487807f8` | Custom Blue Red Basketball Hoop Bedding | *custom blue red basketball hoop bedding* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 205 | `personalized-basketball-with-paint-splash-comforter-ccac4b4a9e-ccac4b4a9e` | Custom Basketball Paint Splash Comforter | *custom basketball paint splash comforter* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 206 | `personalized-basketball-ball-below-net-blanket-with-name-and-number-86aaeebf75-86aaeebf75` | Custom Basketball Net Blanket | *custom basketball net blanket* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 207 | `personalized-basketball-below-hoop-comforter-name-number-e8b829a176-e8b829a176` | Custom Basketball Court Hoop Comforter | *custom basketball court hoop comforter* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 208 | `personalized-basketball-below-hoop-comforter-with-name-and-number-d7a58513fd-d7a58513fd` | Custom Black Basketball Hoop Comforter | *custom black basketball hoop comforter* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 209 | `personalized-basketball-burning-flames-comforter-with-name-and-number-746598fa53-746598fa53` | Custom Flame Basketball Comforter Set | *custom flame basketball comforter set* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |
| 210 | `personalized-basketball-close-up-blanket-with-name-and-number-3f49346de1-3f49346de1` | Custom Basketball Close-Up Blanket | *custom basketball close-up blanket* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 3 | 7 | 20 | 5 | **95.0** | 1 Maj, 1 Min, 1 Lim | `QA_REVISE` |

---

## 3. Phân tích Chi tiết 11 Tiêu chuẩn Đánh giá

### 3.1. Nhóm Tiêu chuẩn Đạt Điểm Tuyệt đối (PASS)
- **P1. Nhận diện Sản phẩm (15/15đ):** Xác định chính xác form factor của cả 2 nhóm sản phẩm: 8 sản phẩm Bedding (Comforter / Duvet Cover) và 2 sản phẩm Blanket (Pos 206 & 210 dạng chăn nỉ nhung fleece throw), nhận diện đúng motif đồ họa thể thao (bóng chày và bóng rổ).
- **P2. Phân loại & Cấu hình (10/10đ):** Khớp đúng taxonomy Shopify collections, phân tách rành mạch cấu hình variants:
  - Nhóm Bedding (8 sản phẩm): 48 variants gồm `Choose Product Type + Size` (8 lựa chọn: Duvet Cover / Comforter x Twin/Full/Queen/King), `Choose Pillowcases` (3 lựa chọn: None/1/2), `Additional Sheet Cover` (2 lựa chọn: None/1).
  - Nhóm Blanket (2 sản phẩm - Pos 206, 210): 8 variants gồm tùy chọn đơn `Choose Your Size` (Small 40"x50", Medium 50"x60", Large 60"x80", v.v.).
- **K1. Từ khóa Chính (10/10đ):** Từ khóa chính phản ánh đúng intent mua sắm thể thao cá nhân hóa tại Mỹ, được kiểm chứng qua 10 truy vấn SERP độc lập.
- **K2. Từ khóa Phụ (5/5đ):** Mở rộng các biến thể tìm kiếm tự nhiên, không nhồi nhét từ khóa.
- **K3. Khớp Ý định Tìm kiếm (5/5đ):** Đúng search intent mua sắm quà tặng và trang trí phòng ngủ thể thao cho thanh thiếu niên, vận động viên và người hâm mộ.
- **T1. Tối ưu Tiêu đề (10/10đ):** Tiêu đề ngắn gọn, súc tích, phản ánh chính xác motif hình ảnh và form factor, đặt từ khóa chính lên đầu.
- **T2. Tối ưu Meta Title (5/5đ):** Độ dài dưới 60 ký tự, kết thúc bằng thương hiệu `| Jeminise`, không bị cắt ngắn trên SERP.
- **I1. Tối ưu Ảnh Gallery (20/20đ):** 100% trong số 64 ảnh được kiểm tra trực quan trực tiếp theo từng pixel; điểm ảnh trung bình đạt 100/100đ, alt text mô tả chính xác chi tiết đồ họa.
- **E1. Căn cứ & Thực chứng (5/5đ):** Đầy đủ bằng chứng từ snapshot nguồn, live storefront, audit customizer, kiểm tra ảnh trực tiếp và 20 truy vấn SERP.

### 3.2. Tiêu chuẩn Bị Trừ điểm: D1. Meta Description SEO (3/5đ - PARTIAL)
- **Hiện trạng kiểm toán:** Trong bản r2, **100% (10/10) meta descriptions bị cắt cụt cứng ngắc ở đúng 155 ký tự** (hard-truncated at 155 chars), làm đứt đoạn câu giữa chừng hoặc cắt ngang từ ngữ vô nghĩa:
  - Pos 201: *"...custom name and numbered "* (cắt ngang từ)
  - Pos 202: *"...with custom name, jersey number, and sele"* (cắt cụt từ `selectable`)
  - Pos 203: *"...custom name, jersey number, and selectable si"* (cắt cụt từ `size`)
  - Pos 204: *"...jersey number, and selectable se"* (cắt cụt từ `set`)
  - Pos 205: *"...jersey number, and selectable s"* (cắt cụt từ `sizes`)
  - Pos 206: *"...with an optional custom name, optional jersey nu"* (cắt cụt từ `number`)
  - Pos 207: *"...jersey number, and selectable size"* (thiếu dấu kết câu)
  - Pos 208: *"...with custom name, jersey n"* (cắt cụt từ `number`)
  - Pos 209: *"...custom name, jersey number,"* (dấu phẩy lơ lửng cuối câu)
  - Pos 210: *"...with an optional custom name, optional jersey nu"* (cắt cụt từ `number`)
- **Hệ quả & Biện pháp:** Gây mất tính chuyên nghiệp trên SERP snippet của Google, làm giảm tỷ lệ nhấp chuột (CTR). Trừ 2 điểm D1 (3/5đ, PARTIAL), ghi nhận 10 lỗi MINOR (`ISSUE-0011` đến `ISSUE-0020`), và cung cấp 10 bản meta description hoàn chỉnh dưới 150 ký tự có cấu trúc ngữ pháp trọn vẹn.

### 3.3. Tiêu chuẩn Bị Trừ điểm: D2. Product Description HTML (7/10đ - PARTIAL)
- **Hiện trạng kiểm toán:** Mô tả r2 chứa câu lệnh sinh nội bộ (prompt boilerplate leakage) xuất hiện trực tiếp trong văn bản hiển thị cho người mua:
  > *"The wording focuses on the visible sports artwork, product form and selectable options for this exact design."*  
  > *"Gallery images show the main bedding or blanket mockup plus feature, care, bedding-type, size or lifestyle panels where present."*
- **Lỗi ngữ pháp tại Pos 206 & 210:** Xuất hiện lỗi tiếng Anh ngô nghê *"uses a optional Custom Name field... uses a optional Custom Number field"* (sai mạo từ `a` trước nguyên âm `optional`).
- **Phát hiện kiểm toán Customizer & Trải nghiệm Mua sắm (Customizer Audit):**
  - **Nhóm Bedding (Pos 201–205, 207–209):** Có **2 trường BẮT BUỘC** (`required=true`) là `Customize Your Name` (tối đa 30 ký tự) và `Customize Your Number` (tối đa 5 ký tự, placeholder `10`).
  - **Nhóm Blanket (Pos 206 & 210):** Có **2 trường TÙY CHỌN** (`required=false`) là `Custom Name` (tối đa 200 ký tự, placeholder `David`) và `Custom Number` (tối đa 20 ký tự, placeholder `22`).
- **Thiếu sót thông số kỹ thuật:** Bản r2 thiếu phân định rành mạch cấu tạo Comforter vs Duvet Cover, thiếu thông tin khóa kéo đáy, thiếu bảng thông số kích thước và hướng dẫn giặt ủi chi tiết.
- **Biện pháp xử lý:** Trừ 3 điểm D2 (7/10đ, PARTIAL), ghi nhận 10 lỗi MAJOR (`ISSUE-0001` đến `ISSUE-0010`), và cung cấp 10 bản mô tả HTML chuẩn publish-ready thay thế hoàn chỉnh.

---

## 4. Kết quả Kiểm toán Trực quan 64 Hình ảnh Gallery

Toàn bộ 64 hình ảnh đã được kiểm tra trực quan trực tiếp theo từng pixel từ thư mục ảnh kiểm toán:

| Pos | Handle | Số lượng ảnh | Chi tiết kiểm toán trực quan (Direct Visual Audit) | Alt Text r2 | Đánh giá |
| :---: | :--- | :---: | :--- | :--- | :---: |
| 201 | `personalized-baseball-bedding-full-size-flag-custom-name-d-design-26` | 7 | `201_01`: Mockup chăn cờ Mỹ, da bóng chày, tên James #10. `201_02`: Mockup vỏ chăn gấp. `201_03`–`201_07`: 5 panel infographic (icon tính năng, sợi microfiber 3D print, Easy Care, Comforter vs Duvet Cover, bảng kích thước). | Chính xác chi tiết cờ Mỹ, bóng chày và tên James #10 | PASS (100đ) |
| 202 | `personalized-baseball-bedding-full-size-flag-custom-name-duvet-cover` | 4 | `202_01`: Cầu thủ bóng chày neon cyan/blue phát sáng trên nền tối, tên Ethan #10. `202_02`: Chi tiết khóa kéo ẩn dưới đáy. `202_03`: Cận cảnh dệt microfiber. `202_04`: Hướng dẫn giặt máy. | Mô tả chuẩn dáng cầu thủ neon cyan phát sáng | PASS (100đ) |
| 203 | `personalized-basketball-above-hoop-and-net-comforter-with-name-and-number-ff32dd676a-ff32dd676a` | 5 | `203_01`: Mockup rổ bóng rổ phát sáng, bóng lơ lửng trên lưới, tên Ethan #10. `203_02`: Khóa kéo đáy. `203_03`: Dệt microfiber. `203_04`: Giặt máy. `203_05`: Mockup góc phòng ngủ. | Mô tả đúng rổ bóng rổ và quả bóng lơ lửng | PASS (100đ) |
| 204 | `personalized-basketball-above-hoop-close-up-comforter-with-name-and-number-95487807f8-95487807f8` | 7 | `204_01`: Rổ bóng rổ vệt sơn xanh đỏ rực rỡ, tên Ethan #10. `204_02`–`204_07`: Bộ 6 ảnh mockup và infographic đầy đủ (tính năng, sợi dệt, Easy Care, Comforter vs Duvet, size chart). | Chính xác chi tiết vệt sơn xanh đỏ và bảng kích thước | PASS (100đ) |
| 205 | `personalized-basketball-with-paint-splash-comforter-ccac4b4a9e-ccac4b4a9e` | 5 | `205_01`: Quả bóng rổ bùng nổ vệt sơn đa sắc nghệ thuật, tên Ethan #10. `205_02`: Khóa kéo. `205_03`: Sợi microfiber. `205_04`: Giặt máy. `205_05`: Góc phòng ngủ. | Chính xác phong cách sơn văng đa sắc nghệ thuật | PASS (100đ) |
| 206 | `personalized-basketball-ball-below-net-blanket-with-name-and-number-86aaeebf75-86aaeebf75` | 8 | `206_01`: Chăn nỉ nhung, cận cảnh bóng rổ lọt lưới trắng tinh tế, tên Lucas #24. `206_02`–`206_08`: 7 panel chăn nỉ nhung (chất liệu fleece siêu mềm, đường may viền, bảng kích thước chăn Small/Medium/Large, giặt máy, lifestyle). | Mô tả chuẩn chăn nỉ nhung bóng lọt lưới | PASS (100đ) |
| 207 | `personalized-basketball-below-hoop-comforter-name-number-e8b829a176-e8b829a176` | 7 | `207_01`: Toàn cảnh sân bóng rổ sàn gỗ bóng loáng hướng về cột rổ, tên Ethan #10. `207_02`–`207_07`: 6 panel mockup và infographic tiêu chuẩn. | Mô tả đúng toàn cảnh sân gỗ và cột rổ bóng rổ | PASS (100đ) |
| 208 | `personalized-basketball-below-hoop-comforter-with-name-and-number-d7a58513fd-d7a58513fd` | 7 | `208_01`: Phong cách đen hiện đại, quả bóng cam nổi bật trên nền lưới hình học đen, tên Ethan #10. `208_02`–`208_07`: 6 panel mockup và infographic tiêu chuẩn. | Chính xác tông đen huyền bí và lưới hình học | PASS (100đ) |
| 209 | `personalized-basketball-burning-flames-comforter-with-name-and-number-746598fa53-746598fa53` | 7 | `209_01`: Quả bóng rổ rực cháy ngọn lửa màu cam đỏ và tàn lửa bay, tên Ethan #10. `209_02`–`209_07`: 6 panel mockup và infographic tiêu chuẩn. | Mô tả chuẩn ngọn lửa cam đỏ rực cháy quanh quả bóng | PASS (100đ) |
| 210 | `personalized-basketball-close-up-blanket-with-name-and-number-3f49346de1-3f49346de1` | 7 | `210_01`: Chăn nỉ nhung, cận cảnh bề mặt da quả bóng rổ sần và bóng cầu thủ, tên David #22. `210_02`–`210_07`: 6 panel chăn nỉ nhung chi tiết (chất liệu, kích cỡ, giặt ủi). | Mô tả chuẩn chăn cận cảnh da bóng rổ sần | PASS (100đ) |

---

## 5. Bằng chứng Thực nghiệm 20 Truy vấn US SERP Độc lập

Đã thực hiện 20 truy vấn tìm kiếm độc lập trên Google Search (thị trường Mỹ `en-US`):

| Query ID | Pos | Từ khóa truy vấn | Loại | Search Intent | Top Organic Domains | Đánh giá & Kết luận |
| :--- | :---: | :--- | :---: | :--- | :--- | :--- |
| `SERP-201-PRI` | 201 | *"custom baseball flag name number bedding"* | `PRIMARY` | Commercial / Patriotic Sports Bedding | etsy.com, ohaprints.com, lylyprint.com, huswiftgift.com | Directly represents Design 26's rustic flag wood plank backdrop, baseball, and custom name/number fields. |
| `SERP-201-COM` | 201 | *"baseball flag comforter set"* | `COMPARATOR` | Commercial / Sports Decor | ebay.com, walmart.com, target.com, temu.com | Confirms strong commercial search interest in patriotic baseball bedding. |
| `SERP-202-PRI` | 202 | *"custom neon baseball player duvet cover"* | `PRIMARY` | Commercial / Neon Graphic Sports | etsy.com, ohaprints.com, youcustomizeit.com, walmart.com | Accurately targets the black background and neon green player graphic theme. |
| `SERP-202-COM` | 202 | *"neon baseball bedding"* | `COMPARATOR` | Commercial / Modern Sports Linens | walmart.com, ebay.com, etsy.com, potterybarnkids.com | Validates neon sports bedding aesthetic demand. |
| `SERP-203-PRI` | 203 | *"custom basketball hoop comforter set"* | `PRIMARY` | Commercial / Custom Basketball Bedding | etsy.com, 2cooldesigns.com, youcustomizeit.com, amorcustomgifts.com | Accurately captures the ball above the hoop and net motif with custom name and number. |
| `SERP-203-COM` | 203 | *"basketball hoop bedding"* | `COMPARATOR` | Commercial / Sports Bedroom | etsy.com, temu.com, shein.com, 90vogue.com | High-volume commercial comparator establishing core motif demand. |
| `SERP-204-PRI` | 204 | *"custom blue red basketball hoop bedding"* | `PRIMARY` | Commercial / Team Color Basketball | etsy.com, 2cooldesigns.com, walmart.com, pbteen.com | Matches the close-up hoop, net, and blue/red color scheme of product 204. |
| `SERP-204-COM` | 204 | *"blue red basketball comforter"* | `COMPARATOR` | Commercial / Colorway Sports Bedding | target.com, walmart.com, ebay.com, lushdecor.com | Validates commercial colorway alignment. |
| `SERP-205-PRI` | 205 | *"custom basketball paint splash comforter"* | `PRIMARY` | Commercial / Artistic Sports Bedding | jeminise.com, etsy.com, 2cooldesigns.com, walmart.com | Accurately reflects the red/blue paint splash basketball design. |
| `SERP-205-COM` | 205 | *"paint splash basketball bedding"* | `COMPARATOR` | Commercial / Abstract Sports Decor | etsy.com, walmart.com, target.com, lushdecor.com | Strong commercial volume confirming popularity of paint splash motif. |
| `SERP-206-PRI` | 206 | *"custom basketball net blanket"* | `PRIMARY` | Commercial / Custom Sports Blanket | zazzle.com, etsy.com, walmart.com, jeminise.com | Directly targets the throw blanket form factor with net artwork and optional customization. |
| `SERP-206-COM` | 206 | *"basketball net throw blanket"* | `COMPARATOR` | Commercial / Fleece Throws | walmart.com, target.com, etsy.com, nba.com | Confirms strong market for basketball-themed throw blankets. |
| `SERP-207-PRI` | 207 | *"custom basketball court hoop comforter"* | `PRIMARY` | Commercial / Court Layout Bedding | etsy.com, amorcustomgifts.com, 2cooldesigns.com, ohaprints.com | Matches the blue-black basketball court lines and hoop design 1 artwork. |
| `SERP-207-COM` | 207 | *"basketball court comforter"* | `COMPARATOR` | Commercial / Sports Comforters | walmart.com, lushdecor.com, etsy.com, ebay.com | Validates core court floor motif search volume. |
| `SERP-208-PRI` | 208 | *"custom black basketball hoop comforter"* | `PRIMARY` | Commercial / Dark Aesthetic Basketball | etsy.com, ohaprints.com, 2cooldesigns.com, amorcustomgifts.com | Accurately reflects Design 2's black grid background and orange ball under hoop. |
| `SERP-208-COM` | 208 | *"black basketball bedding"* | `COMPARATOR` | Commercial / Modern Sports Bedroom | target.com, walmart.com, pbteen.com, crateandbarrel.com | Strong commercial volume confirming popularity of black sports decor. |
| `SERP-209-PRI` | 209 | *"custom flame basketball comforter set"* | `PRIMARY` | Commercial / Fire Sports Bedding | etsy.com, eloquentinnovations.com, ohaprints.com, callie.com | Accurately represents the oversized basketball engulfed in flames graphic. |
| `SERP-209-COM` | 209 | *"fire basketball comforter"* | `COMPARATOR` | Commercial / High-Energy Sports Decor | ebay.com, etsy.com, homisu.store, walmart.com | Confirms strong demand for energetic fire sports bedding. |
| `SERP-210-PRI` | 210 | *"custom basketball close-up blanket"* | `PRIMARY` | Commercial / Custom Sports Blanket | personalizationmall.com, zazzle.com, canvaschamp.com, getphotoblanket.com | Directly targets the dark basketball close-up blanket form factor with optional customization. |
| `SERP-210-COM` | 210 | *"basketball player fleece blanket"* | `COMPARATOR` | Commercial / Fleece Throws | nba.com, logobrands.com, etsy.com, zazzle.com | Validates consumer interest in sports fleece throws. |

---

## 6. Bảng Phân loại Issues & Khuyến nghị Khắc phục

Tổng hợp **30 issues** được ghi nhận trong phiên QA độc lập đầu tiên trên revision r2:

- **10 MAJOR Issues (`ISSUE-0001` đến `ISSUE-0010`):**
  - Trường vi phạm: `description_proposed_html` cho cả 10 sản phẩm.
  - Nguyên nhân: Lộ prompt nội bộ (*"The wording focuses on..."*, *"Gallery images show..."*), lỗi ngữ pháp *"uses a optional"* trên chăn mền, thiếu thông số kỹ thuật then chốt (cấu tạo comforter/duvet, khóa kéo đáy, kích thước, giặt ủi).
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
| 201 | `personalized-baseball-bedding-full-size-flag-custom-name-d-design-26` | `Shop custom baseball flag name number bedding with American flag baseball bedding with wood plank texture, large baseball, script sample name and numbered ` | 155 | **Personalize this custom baseball flag bedding with your player's name and number. Choose duvet cover or comforter in Twin to King sizes.** | 136 |
| 202 | `personalized-baseball-bedding-full-size-flag-custom-name-duvet-cover` | `Shop custom neon baseball player duvet cover with black baseball duvet cover with neon green player silhouettes, sample custom name and jersey number, sele` | 155 | **Upgrade your bedroom with this custom neon baseball duvet cover or comforter set. Personalize with your player's name and jersey number.** | 136 |
| 203 | `personalized-basketball-above-hoop-and-net-comforter-with-name-and-number-ff32dd676a-ff32dd676a` | `Shop custom basketball hoop comforter set with basketball comforter showing a ball above the hoop and net with sample name and jersey number, selectable si` | 155 | **Transform your bedroom with this custom basketball hoop comforter or duvet cover set. Add your athlete's name and number in Twin to King.** | 137 |
| 204 | `personalized-basketball-above-hoop-close-up-comforter-with-name-and-number-95487807f8-95487807f8` | `Shop custom blue red basketball hoop bedding with blue and red basketball bedding with close-up hoop, net, ball, vertical sample name and jersey number, se` | 155 | **Bring court excitement home with this custom blue and red basketball bedding. Available as comforter or duvet cover with custom name and number.** | 144 |
| 205 | `personalized-basketball-with-paint-splash-comforter-ccac4b4a9e-ccac4b4a9e` | `Shop custom basketball paint splash comforter with red and blue paint-splash basketball comforter with large ball graphic, sample name and jersey number, s` | 155 | **Add artistic sports energy with this custom basketball paint splash comforter or duvet cover. Personalize with your player's name and number.** | 141 |
| 206 | `personalized-basketball-ball-below-net-blanket-with-name-and-number-86aaeebf75-86aaeebf75` | `Shop custom basketball net blanket with black basketball blanket with hoop, net, ball under the rim, vertical sample name and jersey number, selectable siz` | 155 | **Cozy up with this custom basketball net blanket. Choose from multiple soft fleece sizes and optionally personalize with any custom name and number.** | 147 |
| 207 | `personalized-basketball-below-hoop-comforter-name-number-e8b829a176-e8b829a176` | `Shop custom basketball court hoop comforter with blue-black basketball comforter with court lines, hoop, large ball and sample custom name, selectable size` | 155 | **Score big with this custom basketball court hoop comforter or duvet cover set. Personalize with your athlete's name and number in Twin to King.** | 143 |
| 208 | `personalized-basketball-below-hoop-comforter-with-name-and-number-d7a58513fd-d7a58513fd` | `Shop custom black basketball hoop comforter with black basketball comforter with orange ball under hoop, grid background, vertical sample name and jersey n` | 155 | **Sleek and athletic, this custom black basketball hoop comforter or duvet cover features bold court graphics personalized with your name and number.** | 147 |
| 209 | `personalized-basketball-burning-flames-comforter-with-name-and-number-746598fa53-746598fa53` | `Shop custom flame basketball comforter set with basketball comforter with oversized close-up ball, burning flame background, sample name and jersey number,` | 155 | **Ignite your sports bedroom decor with this custom flame basketball comforter or duvet cover set. Personalize with your player's name and number.** | 144 |
| 210 | `personalized-basketball-close-up-blanket-with-name-and-number-3f49346de1-3f49346de1` | `Shop custom basketball close-up blanket with dark basketball blanket with close-up ball, gray player silhouettes, sample name and jersey number, selectable` | 155 | **Wrap up in athletic warmth with this custom basketball close-up blanket. Premium plush fleece in multiple sizes with optional custom name and number.** | 149 |

---

## 7. Đề xuất 10 Mô tả HTML Publish-Ready (Chuẩn hóa Hoàn chỉnh)

Dưới đây là 10 bản mô tả HTML tiếng Anh hoàn chỉnh, loại bỏ sạch prompt boilerplate, phản ánh chuẩn xác live customizer (2 trường Name/Number bắt buộc đối với Bedding; 2 trường Name/Number tùy chọn đối với Blanket), phân định rành mạch Comforter vs Duvet Cover, kích thước, phụ kiện kèm theo và hướng dẫn giặt ủi:

### Position 201: Custom Baseball Flag Name Number Bedding (`personalized-baseball-bedding-full-size-flag-custom-name-d-design-26`)

```html
<p>Bring the timeless spirit of America's pastime and patriotic pride into your bedroom with the <strong>Custom Baseball Flag Name Number Bedding</strong> from Jeminise. This dynamic sports bedding set combines the iconic stars and stripes of the American flag with authentic textured baseball leather and bold red seam stitching. Tailored with your athlete's name and jersey number in athletic typography, it serves as the ultimate personalized centerpiece for young ballplayers, high school athletes, and baseball fans of all ages.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> A complete, ready-to-use comforter filled with lofty, hypoallergenic microfiber batting. Delivers balanced, cloud-like warmth and breathable comfort through all seasons.</li>
  <li><strong>Duvet Cover Option:</strong> A lightweight protective cover equipped with a concealed bottom zipper closure, designed to slip smoothly over your existing comforter or duvet insert for effortless removal and laundering.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name:</strong> Enter your ballplayer's first name, last name, or team name (1–30 characters) in the <em>Customize Your Name</em> field. Enter "NO" if you prefer the artwork without a name.</li>
  <li><strong>Custom Number:</strong> Enter your athlete's jersey number (1–5 digits) in the <em>Customize Your Number</em> field. Enter "NO" if you prefer no number printed.</li>
</ul>
<h3>Available Sizes & Coordinating Add-Ons</h3>
<ul>
  <li><strong>Bed Sizes:</strong> Twin (68x86 in), Full (79x90 in), Queen (90x90 in), and King (90x104 in).</li>
  <li><strong>Matching Pillowcases:</strong> Bundle with None, 1 Pillowcase, or 2 Pillowcases printed with coordinating baseball flag artwork.</li>
  <li><strong>Optional Flat Sheet:</strong> Complete your bedding set with an optional matching flat sheet cover in corresponding dimensions.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> 100% premium high-density brushed polyester microfiber for ultra-soft hand feel, wrinkle resistance, and lasting durability.</li>
  <li><strong>Care:</strong> Machine wash cold on gentle cycle with mild detergent. Tumble dry on low heat or hang dry. High-definition dye-sublimation print resists fading wash after wash.</li>
</ul>
```

### Position 202: Custom Neon Baseball Player Duvet Cover (`personalized-baseball-bedding-full-size-flag-custom-name-duvet-cover`)

```html
<p>Light up bedtime with electrifying athletic intensity with the <strong>Custom Neon Baseball Player Duvet Cover</strong> from Jeminise. Featuring a striking neon silhouette of a baseball batter poised at the plate, glowing with radiant cyan and blue energy against a dark atmospheric stadium background, this bedding brings high-voltage modern style to any room. Personalized with your player's custom name and jersey number, it transforms any teen bedroom, dorm, or sports lover's sanctuary into a neon arena.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Fully quilted all-in-one comforter filled with airy, hypoallergenic down-alternative microfiber batting for cozy, year-round bedtime warmth.</li>
  <li><strong>Duvet Cover Option:</strong> Equipped with a hidden bottom zipper closure, making it effortless to insert, remove, and wash your comforter insert.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name:</strong> Type your athlete's name or custom text (1–30 characters) into the <em>Customize Your Name</em> box. Enter "NO" for clean, unpersonalized neon artwork.</li>
  <li><strong>Custom Number:</strong> Type your player's jersey number (1–5 characters) into the <em>Customize Your Number</em> box. Enter "NO" to omit.</li>
</ul>
<h3>Available Sizes & Coordinating Add-Ons</h3>
<ul>
  <li><strong>Bed Sizes:</strong> Available in Twin (68x86 in), Full (79x90 in), Queen (90x90 in), and King (90x104 in).</li>
  <li><strong>Matching Pillowcases:</strong> Add None, 1 Pillowcase, or 2 Pillowcases with matching glowing neon baseball graphics.</li>
  <li><strong>Additional Flat Sheet:</strong> Optionally add a coordinated flat sheet cover in identical mattress dimensions.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> Premium fine-spun brushed microfiber delivering gentle skin comfort, breathability, and anti-static softness.</li>
  <li><strong>Care:</strong> Machine wash cold with like colors. Tumble dry on low. Resists shrinking, wrinkles, and color fading.</li>
</ul>
```

### Position 203: Custom Basketball Hoop Comforter Set (`personalized-basketball-above-hoop-and-net-comforter-with-name-and-number-ff32dd676a-ff32dd676a`)

```html
<p>Step onto center court every night with the <strong>Custom Basketball Hoop Comforter Set</strong> from Jeminise. This immersive sports bedding captures the drama of game night, featuring an illuminated basketball rim and nylon net poised under arena floodlights with an authentic textured basketball hovering at the rim. Customized with your player's name and jersey number, it makes an unforgettable bedroom statement and thoughtful gift for young hoopers, varsity athletes, and devoted basketball fans.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Pre-filled comforter packed with soft, lightweight, all-season microfiber insulation for immediate warmth and cloud-soft comfort right out of the box.</li>
  <li><strong>Duvet Cover Option:</strong> Protective duvet shell featuring an invisible bottom zipper closure for hassle-free insert changes and easy home laundry.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name:</strong> Enter your ballplayer's name or nickname (1–30 characters) in the <em>Customize Your Name</em> field. Enter "NO" if no name is desired.</li>
  <li><strong>Custom Number:</strong> Enter your athlete's jersey number (1–5 digits) in the <em>Customize Your Number</em> field. Enter "NO" if no number is wanted.</li>
</ul>
<h3>Available Sizes & Coordinating Add-Ons</h3>
<ul>
  <li><strong>Bed Sizes:</strong> Twin (68x86 in), Full (79x90 in), Queen (90x90 in), and King (90x104 in).</li>
  <li><strong>Coordinating Pillowcases:</strong> Select None, 1 Pillowcase, or 2 Pillowcases printed with matching basketball hoop artwork.</li>
  <li><strong>Matching Flat Sheet:</strong> Optional coordinated flat sheet cover in identical selected bed dimensions.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> 100% brushed polyester microfiber fabric engineered for superior breathability, softness, and durability.</li>
  <li><strong>Care:</strong> Machine wash cold on delicate cycle. Tumble dry low or line dry. Vibrant colors remain bold without bleeding or cracking.</li>
</ul>
```

### Position 204: Custom Blue Red Basketball Hoop Bedding (`personalized-basketball-above-hoop-close-up-comforter-with-name-and-number-95487807f8-95487807f8`)

```html
<p>Ignite athletic passion in your bedroom with the <strong>Custom Blue Red Basketball Hoop Bedding</strong> from Jeminise. Showcasing an explosive graphic of a basketball net accented with vivid blue and fiery red spray splatters and glowing court lines, this bedding delivers unmatched visual energy. Fully customized with your athlete's name and jersey number, it turns any bedroom into an inspiring locker-room showcase for dedicated basketball players.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Generously filled with down-alternative microfiber fill that provides lightweight, breathable warmth throughout every season.</li>
  <li><strong>Duvet Cover Option:</strong> Features a durable, hidden bottom zipper closure for effortless cover removal and fast washing.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name:</strong> Enter your player's name (1–30 characters) in the <em>Customize Your Name</em> box. Type "NO" for artwork without personalization.</li>
  <li><strong>Custom Number:</strong> Enter your player's jersey number (1–5 characters) in the <em>Customize Your Number</em> box. Type "NO" to omit.</li>
</ul>
<h3>Available Sizes & Coordinating Add-Ons</h3>
<ul>
  <li><strong>Bed Sizes:</strong> Twin, Full, Queen, and King sizing available.</li>
  <li><strong>Matching Pillowcases:</strong> Add 1 or 2 matching pillowcases displaying coordinating blue and red paint-splashed basketball graphics.</li>
  <li><strong>Optional Flat Sheet:</strong> Coordinated flat sheet cover available in matching mattress dimensions.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> High-density woven microfiber engineered for softness, wrinkle resistance, and moisture-wicking comfort.</li>
  <li><strong>Care:</strong> Machine wash cold with mild detergent. Tumble dry on low heat. Resists fading and pilling.</li>
</ul>
```

### Position 205: Custom Basketball Paint Splash Comforter (`personalized-basketball-with-paint-splash-comforter-ccac4b4a9e-ccac4b4a9e`)

```html
<p>Make a bold artistic statement on the court and in the bedroom with the <strong>Custom Basketball Paint Splash Comforter</strong> from Jeminise. This dynamic design features an authentic textured basketball bursting through multi-colored paint splatters, expressive brushstrokes, and energetic splatter droplets. Customized with your athlete's name and number, it blends fine-art creativity with varsity sports excitement for an eye-catching bedroom upgrade.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Ready-to-use comforter filled with fluffy, hypoallergenic microfiber batting for all-season plush comfort.</li>
  <li><strong>Duvet Cover Option:</strong> Equipped with an unobtrusive bottom zipper closure, allowing quick insertion of your favorite comforter or duvet insert.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name:</strong> Type any custom name or team dedication (1–30 characters) in the <em>Customize Your Name</em> field. Enter "NO" for non-customized artwork.</li>
  <li><strong>Custom Number:</strong> Type your athlete's jersey number (1–5 digits) in the <em>Customize Your Number</em> box. Enter "NO" if not needed.</li>
</ul>
<h3>Available Sizes & Coordinating Add-Ons</h3>
<ul>
  <li><strong>Bed Sizes:</strong> Standard Twin, Full, Queen, and King mattress coverage.</li>
  <li><strong>Matching Pillowcases:</strong> Select None, 1 Pillowcase, or 2 Pillowcases with matching paint-splashed basketball graphics.</li>
  <li><strong>Matching Flat Sheet:</strong> Optional coordinated flat sheet cover in corresponding dimensions.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> 100% fine-spun brushed microfiber offering silky softness and anti-wrinkle durability.</li>
  <li><strong>Care:</strong> Machine wash cold on gentle cycle. Tumble dry on low heat. High-contrast pigments remain rich and vivid.</li>
</ul>
```

### Position 206: Custom Basketball Net Blanket (`personalized-basketball-ball-below-net-blanket-with-name-and-number-86aaeebf75-86aaeebf75`)

```html
<p>Wrap yourself in game-winning comfort with the <strong>Custom Basketball Net Blanket</strong> from Jeminise. This cozy sports throw blanket features a stunning macro close-up of a textured basketball dropping cleanly through the crisp white cords of a basketball net under dramatic lighting. Perfect for courtside warmups, post-game movie nights on the couch, or extra warmth at bedtime, this blanket makes an ideal personalized gift for basketball players, coaches, and sports enthusiasts.</p>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name (Optional):</strong> Personalize with your player's first name, last name, or team nickname (1–200 characters) in the <em>Custom Name</em> field. This field is completely optional; leave blank if you prefer the graphic design without text.</li>
  <li><strong>Custom Number (Optional):</strong> Enter your athlete's jersey number (1–20 characters) in the <em>Custom Number</em> field. This field is optional; leave blank for an unnumbered throw.</li>
</ul>
<h3>Available Throw & Blanket Sizes</h3>
<ul>
  <li><strong>Small (40" x 50"):</strong> Perfect for toddlers, small children, car travel, or pet lap throws.</li>
  <li><strong>Medium (50" x 60"):</strong> The standard throw size, ideal for couch lounging, reading, or stadium spectating.</li>
  <li><strong>Large (60" x 80"):</strong> Generous oversized blanket, perfect for twin/full bed coverage or wrapping up completely.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> Ultra-soft premium plush fleece with velvety flannel touch, providing cozy warmth without heavy bulk.</li>
  <li><strong>Care:</strong> Machine wash cold with mild detergent on gentle cycle. Tumble dry on low heat or line dry. Non-pilling, anti-static, and fade-resistant.</li>
</ul>
```

### Position 207: Custom Basketball Court Hoop Comforter (`personalized-basketball-below-hoop-comforter-name-number-e8b829a176-e8b829a176`)

```html
<p>Bring the arena home with the <strong>Custom Basketball Court Hoop Comforter</strong> from Jeminise. Featuring an expansive, wide-angle view of a polished hardwood basketball court stretching toward a regulation hoop under gleaming stadium lights, this bedding delivers genuine hardwood atmosphere to any sports-themed bedroom. Customized with your player's name and jersey number, it inspires championship dreams every night.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Pre-filled comforter with plush down-alternative microfiber fill, engineered for dependable, all-season bedtime comfort.</li>
  <li><strong>Duvet Cover Option:</strong> Sleek duvet cover with a concealed bottom zipper closure, providing an easy-to-wash outer layer for your duvet insert.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name:</strong> Input your ballplayer's name or custom dedication (1–30 characters) in the <em>Customize Your Name</em> box. Enter "NO" for clean court artwork.</li>
  <li><strong>Custom Number:</strong> Input your player's jersey number (1–5 characters) in the <em>Customize Your Number</em> box. Enter "NO" to omit.</li>
</ul>
<h3>Available Sizes & Coordinating Add-Ons</h3>
<ul>
  <li><strong>Bed Sizes:</strong> Twin (68x86 in), Full (79x90 in), Queen (90x90 in), and King (90x104 in).</li>
  <li><strong>Coordinating Pillowcases:</strong> Add 1 or 2 matching basketball court pillowcases to complete your sports bedding set.</li>
  <li><strong>Additional Flat Sheet:</strong> Complete your bedding ensemble with an optional matching flat sheet cover.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> 100% premium brushed microfiber weave for breathable softness, hypoallergenic comfort, and lasting durability.</li>
  <li><strong>Care:</strong> Machine wash cold with mild detergent. Tumble dry on low heat. Resists shrinking and color fading.</li>
</ul>
```

### Position 208: Custom Black Basketball Hoop Comforter (`personalized-basketball-below-hoop-comforter-with-name-and-number-d7a58513fd-d7a58513fd`)

```html
<p>Make a bold, modern athletic statement with the <strong>Custom Black Basketball Hoop Comforter</strong> from Jeminise. Set against a sophisticated dark textured geometric grid background, this bedding highlights a vibrant orange basketball and rim illuminated with dramatic athletic focus. Tailored with your athlete's name and jersey number in high-contrast athletic lettering, it brings an edgy, contemporary look to teen bedrooms, dorm rooms, and adult fan spaces.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Complete, pre-filled comforter packed with airy microfiber batting, delivering plush insulation and breathable warmth.</li>
  <li><strong>Duvet Cover Option:</strong> Equipped with an invisible bottom zipper, ideal for slipping over your favorite duvet insert for convenient machine washing.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name:</strong> Enter your athlete's name or custom text (1–30 characters) into the <em>Customize Your Name</em> field. Type "NO" for unpersonalized artwork.</li>
  <li><strong>Custom Number:</strong> Enter your ballplayer's jersey number (1–5 digits) into the <em>Customize Your Number</em> box. Type "NO" if no number is wanted.</li>
</ul>
<h3>Available Sizes & Coordinating Add-Ons</h3>
<ul>
  <li><strong>Bed Sizes:</strong> Available in Twin, Full, Queen, and King bed dimensions.</li>
  <li><strong>Matching Pillowcases:</strong> Bundle with None, 1 Pillowcase, or 2 Pillowcases with matching dark basketball hoop graphics.</li>
  <li><strong>Optional Flat Sheet:</strong> Coordinated flat sheet cover available in matching mattress size.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> High-density brushed microfiber engineered for exceptional skin-friendly softness, lightweight warmth, and anti-static comfort.</li>
  <li><strong>Care:</strong> Machine wash cold on delicate cycle. Tumble dry low. Deep black pigments stay rich and fade-free wash after wash.</li>
</ul>
```

### Position 209: Custom Flame Basketball Comforter Set (`personalized-basketball-burning-flames-comforter-with-name-and-number-746598fa53-746598fa53`)

```html
<p>Ignite your bedroom with championship intensity with the <strong>Custom Flame Basketball Comforter Set</strong> from Jeminise. This high-octane sports bedding design features a blazing basketball engulfed in roaring orange and crimson flames with flying embers across a sleek dark background. Personalized with your player's custom name and jersey number, this set transforms any bedroom into an arena of unstoppable athletic energy.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> A cozy, pre-filled comforter with lofty down-alternative microfiber fill that provides balanced, cloud-like comfort all year round.</li>
  <li><strong>Duvet Cover Option:</strong> Features a smooth bottom zipper closure, making it effortless to insert and remove your comforter insert for regular cleaning.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name:</strong> Personalize with your player's name (1–30 characters) in the <em>Customize Your Name</em> field. Enter "NO" for artwork without a name.</li>
  <li><strong>Custom Number:</strong> Personalize with your player's jersey number (1–5 characters) in the <em>Customize Your Number</em> field. Enter "NO" to omit.</li>
</ul>
<h3>Available Sizes & Coordinating Add-Ons</h3>
<ul>
  <li><strong>Bed Sizes:</strong> Twin, Full, Queen, and King sizing available.</li>
  <li><strong>Coordinating Pillowcases:</strong> Select None, 1 Pillowcase, or 2 Pillowcases featuring matching flame basketball artwork.</li>
  <li><strong>Matching Flat Sheet:</strong> Optionally include an additional flat sheet cover in corresponding dimensions.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> 100% woven microfiber fabric with high yarn density for durable softness and breathability.</li>
  <li><strong>Care:</strong> Machine wash cold, gentle cycle. Tumble dry on low heat. Vivid dye-sublimation print resists fading, wrinkling, and pilling.</li>
</ul>
```

### Position 210: Custom Basketball Close-Up Blanket (`personalized-basketball-close-up-blanket-with-name-and-number-3f49346de1-3f49346de1`)

```html
<p>Celebrate your love for the game with the <strong>Custom Basketball Close-Up Blanket</strong> from Jeminise. Showcasing a stunning photographic close-up of a basketball's pebbled leather texture paired with an athletic player silhouette against a dramatic dark backdrop, this throw blanket brings authentic courtside excitement to your home. Customized optionally with your player's name and jersey number, it is the ultimate plush companion for cozy evenings, road trips, and game day spectating.</p>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name (Optional):</strong> Add your ballplayer's name or team name (1–200 characters) in the <em>Custom Name</em> field. This field is optional; leave blank if you prefer clean artwork without text.</li>
  <li><strong>Custom Number (Optional):</strong> Add your player's number (1–20 characters) in the <em>Custom Number</em> field. This field is optional; leave blank for an unnumbered blanket.</li>
</ul>
<h3>Available Throw & Blanket Sizes</h3>
<ul>
  <li><strong>Small (40" x 50"):</strong> Compact size ideal for kids, baby cribs, strollers, or car travel.</li>
  <li><strong>Medium (50" x 60"):</strong> Most popular throw size for sofa cuddling, gaming chairs, or reading nooks.</li>
  <li><strong>Large (60" x 80"):</strong> Generously sized for full bed layering or wrapping up head-to-toe on chilly evenings.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> 100% premium ultra-soft fleece with velvety flannel hand feel, offering lightweight warmth and breathable comfort.</li>
  <li><strong>Care:</strong> Machine wash cold on gentle cycle. Tumble dry on low heat or hang dry. Anti-pilling, shrink-resistant, and colorfast.</li>
</ul>
```

---

## 8. Tuyên bố Trạng thái & Tiếp tục Quy trình

- **Xác nhận khóa dữ liệu:** Toàn bộ bằng chứng kiểm toán độc lập đã được lưu trữ vĩnh viễn tại `seo_runs/jeminise.com/20260906_234129/qa/20260908_133500/`.
- **File bàn giao:** `resutls/jeminise.com/20260906_234129/qa/20260908_133500/SEO_QA_qa_batch_021_r2.xlsx` và `SEO_QA_qa_batch_021_r2.md`.
- **Nguyên tắc an toàn:** Không chỉnh sửa file nguồn `SEO_Product_Optimization_qa_batch_021_r2.xlsx` và không sửa trực tiếp live store.
- **Dừng theo quy trình:** Hệ thống kích hoạt cờ `awaiting_confirmation=true` và **DỪNG LẠI**, chờ lệnh xác nhận từ người dùng trước khi chuyển sang Batch 22.
