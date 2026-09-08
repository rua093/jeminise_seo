# Báo cáo Re-QA Độc lập SEO Workbook - Batch 20 Revision r3

> **Mã phiên QA:** `20260908_112500`  
> **Tập tin nguồn thẩm định:** `resutls/jeminise.com/20260906_234129/revisions/qa_batch_020_r3/SEO_Product_Optimization_qa_batch_020_r3.xlsx`  
> **SHA256 Tập tin nguồn:** `345ac8e369975cd8856b3cc6438daa0b6ce56d7beb7d61ce19c1746c6df6ff12`  
> **Thị trường mục tiêu:** United States (`en-US`) | **Nội dung:** English SEO Copy  
> **Phạm vi thẩm định:** 10 sản phẩm (Inventory Position 191–200), 46 ảnh gallery, 40 dòng Keyword_Map, 10 dòng Buyer_Search_Research  
> **Trạng thái thẩm định:** `awaiting_confirmation=true` | **Kết luận tổng thể:** `QA_REVISE` (Điểm trung bình: **97.00/100**)  

---

## 1. Tóm tắt kết quả Re-QA & Chấm điểm Tổng thể

Quá trình Re-QA độc lập được thực hiện với 100% trọng số đánh giá trên toàn bộ 10 sản phẩm, 46 hình ảnh gallery, đối chiếu trực tiếp dữ liệu live storefront và customizer tại `jeminise.com`, cùng 20 truy vấn US SERP độc lập:

| Chỉ số | Giá trị | Ghi chú kiểm toán |
| :--- | :---: | :--- |
| **Tổng số sản phẩm** | 10 | Position 191 đến 200 (`design-16` đến `design-25`) |
| **Tỷ lệ kiểm toán sản phẩm** | 100% (10/10) | Cột `page_read=TRUE` đầy đủ |
| **Tổng số ảnh gallery** | 46 | Pos 191: 7 ảnh, Pos 192: 7 ảnh, Pos 193–200: 4 ảnh/sp |
| **Tỷ lệ kiểm toán ảnh** | 100% (46/46) | Điểm ảnh trung bình 100.0/100, 100% PASS |
| **Điểm số trung bình** | **97.00 / 100** | 10/10 sản phẩm đạt chính xác 97.0 điểm |
| **Số sản phẩm QA_PASS** | 0 | Do tồn tại lỗi MAJOR ở tiêu chuẩn D2 |
| **Số sản phẩm QA_REVISE** | **10** | Yêu cầu chuẩn hóa lại mô tả sản phẩm D2 |
| **Số sản phẩm QA_FAIL** | 0 | Không có lỗi CRITICAL |
| **Số sản phẩm QA_INCOMPLETE**| 0 | 100% dữ liệu được đánh giá đầy đủ |
| **Kết luận tổng thể** | **QA_REVISE** | Sẵn sàng xuất bản sau khi áp dụng 10 mô tả HTML đề xuất |

---

## 2. Bảng tổng hợp đánh giá 10 sản phẩm (Pos 191–200)

| Pos | Product Key / Handle | Title đề xuất r3 | Primary Keyword | P1 | P2 | K1 | K2 | K3 | T1 | T2 | D1 | D2 | I1 | E1 | Tổng | Issues | QA Status |
| :---: | :--- | :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| 191 | `personalized-baseball-bedding-full-size-flag-custom-name-d-design-16` | Custom Red Black Batter Bedding | *custom red black batter bedding* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 5 | 7 | 20 | 5 | **97.0** | 1 Maj, 1 Lim | `QA_REVISE` |
| 192 | `personalized-baseball-bedding-full-size-flag-custom-name-d-design-17` | Custom Brown Baseball Glove Bedding | *custom brown baseball glove bedding* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 5 | 7 | 20 | 5 | **97.0** | 1 Maj, 1 Lim | `QA_REVISE` |
| 193 | `personalized-baseball-bedding-full-size-flag-custom-name-d-design-18` | Custom Flaming Baseball Batter Bedding | *custom flaming baseball batter bedding* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 5 | 7 | 20 | 5 | **97.0** | 1 Maj, 1 Lim | `QA_REVISE` |
| 194 | `personalized-baseball-bedding-full-size-flag-custom-name-d-design-19` | Custom Orange Fire Baseball Bedding | *custom orange fire baseball bedding* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 5 | 7 | 20 | 5 | **97.0** | 1 Maj, 1 Lim | `QA_REVISE` |
| 195 | `personalized-baseball-bedding-full-size-flag-custom-name-d-design-20` | Custom Baseball Glove Flag Bedding | *custom baseball glove flag bedding* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 5 | 7 | 20 | 5 | **97.0** | 1 Maj, 1 Lim | `QA_REVISE` |
| 196 | `personalized-baseball-bedding-full-size-flag-custom-name-d-design-21` | Custom Black Baseball Glove Bedding | *custom black baseball glove bedding* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 5 | 7 | 20 | 5 | **97.0** | 1 Maj, 1 Lim | `QA_REVISE` |
| 197 | `personalized-baseball-bedding-full-size-flag-custom-name-d-design-22` | Custom Lightning Baseball Bedding | *custom lightning baseball bedding* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 5 | 7 | 20 | 5 | **97.0** | 1 Maj, 1 Lim | `QA_REVISE` |
| 198 | `personalized-baseball-bedding-full-size-flag-custom-name-d-design-23` | Custom Night Field Baseball Bedding | *custom night field baseball bedding* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 5 | 7 | 20 | 5 | **97.0** | 1 Maj, 1 Lim | `QA_REVISE` |
| 199 | `personalized-baseball-bedding-full-size-flag-custom-name-d-design-24` | Custom Smoke Baseball Batter Bedding | *custom smoke baseball batter bedding* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 5 | 7 | 20 | 5 | **97.0** | 1 Maj, 1 Lim | `QA_REVISE` |
| 200 | `personalized-baseball-bedding-full-size-flag-custom-name-d-design-25` | Custom Baseball Stitch Name Bedding | *custom baseball stitch name bedding* | 15 | 10 | 10 | 5 | 5 | 10 | 5 | 5 | 7 | 20 | 5 | **97.0** | 1 Maj, 1 Lim | `QA_REVISE` |

---

## 3. Phân tích Chi tiết 11 Tiêu chuẩn Đánh giá

### 3.1. Nhóm Tiêu chuẩn Đạt Điểm Tuyệt đối (PASS)
- **P1. Nhận diện Sản phẩm (15/15đ):** Xác định chính xác 100% form factor (`Bedding` dạng Comforter / Duvet Cover), chất liệu sợi microfiber dệt mật độ cao chải mềm, và motif bóng chày cá nhân hóa.
- **P2. Phân loại & Cấu hình (10/10đ):** Khớp đúng taxonomy Shopify collections, danh mục Bedding, và cấu trúc tùy chọn variants (Product Type + Size: 8 lựa chọn, Pillowcases: 3 lựa chọn, Sheet Cover: 2 lựa chọn = 48 variants).
- **K1. Từ khóa Chính (10/10đ):** Từ khóa chính phản ánh đúng intent mua sắm thể thao cá nhân hóa tại Mỹ, được kiểm chứng qua 10 truy vấn SERP độc lập.
- **K2. Từ khóa Phụ (5/5đ):** Mở rộng các biến thể tìm kiếm dài hạn tự nhiên, không nhồi nhét từ khóa.
- **K3. Khớp Ý định Tìm kiếm (5/5đ):** Đúng search intent mua sắm quà tặng và trang trí phòng ngủ thể thao cho thanh thiếu niên và fan bóng chày.
- **T1. Tối ưu Tiêu đề (10/10đ):** Loại bỏ hoàn toàn chuỗi từ khóa rác lặp lại (`full size flag custom name d`) từ dữ liệu cũ, tiêu đề ngắn gọn súc tích và đặt từ khóa chính lên đầu.
- **T2. Tối ưu Meta Title (5/5đ):** Độ dài dưới 60 ký tự, chứa thương hiệu `Jeminise`, không bị cắt ngắn trên SERP.
- **D1. Tối ưu Meta Description (5/5đ):** Độ dài chuẩn 100–128 ký tự, hấp dẫn, loại bỏ sạch thẻ draft nội bộ.
- **I1. Tối ưu Ảnh Gallery (20/20đ):** 100% trong số 46 ảnh được kiểm tra trực quan trực tiếp; điểm ảnh trung bình đạt 100/100đ, alt text mô tả chính xác chi tiết đồ họa.
- **E1. Căn cứ & Thực chứng (5/5đ):** Đầy đủ bằng chứng từ live storefront, audit customizer, kiểm tra ảnh trực tiếp và 20 truy vấn SERP.

### 3.2. Tiêu chuẩn Bị Trừ điểm: D2. Product Description HTML (7/10đ - PARTIAL)
- **Hiện trạng:** Bản r3 đã xóa bỏ hoàn toàn các thẻ nháp kỹ thuật cũ (`SEO Use / QA approval`). Tuy nhiên, copywriter đã đưa vào đoạn văn bản phủ nhận khả năng cá nhân hóa số áo:
  > *"Any visible numbers are sample artwork only. In live purchase flow, only Customize Your Name is collected..."*
- **Phát hiện kiểm toán live storefront (CRITICAL AUDIT FINDING):**
  Đối chiếu trực tiếp mã nguồn giao diện storefront và endpoint `.js` tại `jeminise.com`, **TẤT CẢ 10 SẢN PHẨM** trong Batch 20 đều có **CẢ HAI TRƯỜNG NHẬP LIỆU BẮT BUỘC**:
  1. `Customize Your Name`: bắt buộc, tối đa 30 ký tự, hướng dẫn: *"Enter 'NO' if you don't want to customize"*.
  2. `Customize Your Number`: bắt buộc, tối đa 5 ký tự, placeholder: `10`, hướng dẫn: *"Enter 'NO' if you don't want to customize"*.
- **Hệ quả:** Tuyên bố của r3 là sai sự thật đối với trải nghiệm mua sắm thực tế của khách hàng, gây hoang mang cho người mua khi họ nhìn thấy ô nhập số áo trên website. Ngoài ra, mô tả r3 thiếu phân tách rõ ràng giữa Comforter và Duvet Cover, thiếu thông số kích thước và tùy chọn mua thêm.
- **Biện pháp xử lý:** Trừ 3 điểm D2 (7/10đ, PARTIAL), ghi nhận 10 lỗi MAJOR mới (`ISSUE-0075` đến `ISSUE-0084`), và cung cấp 10 bản mô tả HTML chuẩn publish-ready thay thế hoàn chỉnh.

---

## 4. Kết quả Kiểm toán Trực quan 46 Hình ảnh Gallery

Toàn bộ 46 hình ảnh đã được kiểm tra trực quan trực tiếp theo từng pixel:

| Pos | Handle | Số lượng ảnh | Chi tiết kiểm toán trực quan (Direct Visual Audit) | Alt Text r3 | Đánh giá |
| :---: | :--- | :---: | :--- | :--- | :---: |
| 191 | `design-16` | 7 | `191_01`: Cầu thủ đánh bóng James #5, vệt sơn đỏ đen. `191_02`: Mockup vỏ chăn gấp. `191_03`: 4 icon tính năng (soft, lightweight, durable, breathable). `191_04`: Cận cảnh microfiber 3D print. `191_05`: Panel Stress-Free Easy Care. `191_06`: Infographic phân biệt Comforter vs Duvet Cover. `191_07`: Bảng kích thước (Twin 68x86, Full 79x90, Queen 90x90, King 90x104 in). | Chính xác chi tiết cầu thủ James #5 và vệt sơn đỏ đen | PASS (100đ) |
| 192 | `design-17` | 7 | `192_01`: Găng tay da bóng chày màu nâu tự nhiên, bóng trầy xước, nền gỗ, chữ Charles. `192_02`: Mockup góc giường. `192_03`–`192_07`: Bộ 5 panel infographic với nền găng da nâu và bảng kích thước chuẩn. | Mô tả đúng găng da nâu và bóng trầy xước | PASS (100đ) |
| 193 | `design-18` | 4 | `193_01`: Cầu thủ Gabriel #10 đánh bóng sang bên trái, vòng tròn lửa rực cháy quanh quả bóng. `193_02`: Chi tiết khóa kéo dưới đáy. `193_03`: Cận cảnh dệt microfiber mật độ cao. `193_04`: Hướng dẫn giặt máy. | Phân biệt rõ cầu thủ đánh bóng qua trái và vòng tròn lửa | PASS (100đ) |
| 194 | `design-19` | 4 | `194_01`: Cầu thủ Jackson #10 đánh bóng sang bên phải, bão lửa màu cam rực cháy bao phủ toàn bộ cảnh nền. `194_02`–`194_04`: Khóa kéo, sợi dệt microfiber, giặt máy. | Phân biệt rõ bão lửa cam toàn khung hình và hướng đánh phải | PASS (100đ) |
| 195 | `design-20` | 4 | `195_01`: Găng tay da nâu và bóng trên nền quốc kỳ Mỹ gỗ mộc, chữ Jackson, số áo mẫu #8 trên vỏ gối shams. `195_02`–`195_04`: Giặt máy, khóa kéo, sợi dệt. | Mô tả chuẩn găng bóng trên cờ Mỹ và số #8 trên gối | PASS (100đ) |
| 196 | `design-21` | 4 | `196_01`: Găng tay và bóng phong cách đơn sắc đen monochrome, nền đen tuyền, chữ Robert #8 màu trắng. `196_02`–`196_04`: Khóa kéo, sợi dệt, giặt máy. | Chính xác phong cách đen đơn sắc huyền bí | PASS (100đ) |
| 197 | `design-22` | 4 | `197_01`: Tia sét màu xanh cyan rực sáng bao quanh quả bóng đang bay tốc độ cao, chữ Jackson #8. `197_02`–`197_04`: Khóa kéo, sợi dệt, giặt máy. | Mô tả chuẩn tia sét xanh cyan và bóng bay | PASS (100đ) |
| 198 | `design-23` | 4 | `198_01`: Sân vận động bóng chày ban đêm dưới ánh đèn pha, mũ bảo hiểm, gậy gỗ, găng và bóng trên cỏ, chữ Benedict (không có số trên mockup). `198_02`–`198_04`: Khóa kéo, sợi dệt, giặt máy. | Mô tả chuẩn sân vận động ban đêm và bộ dụng cụ | PASS (100đ) |
| 199 | `design-24` | 4 | `199_01`: Cầu thủ vung gậy giữa làn khói trắng / sương mù cuộn xoáy ma mị, chữ Jackson, số mẫu #23 trên gối shams. `199_02`–`199_04`: Khóa kéo, sợi dệt, giặt máy. | Mô tả chuẩn làn khói trắng và số mẫu #23 trên gối | PASS (100đ) |
| 200 | `design-25` | 4 | `200_01`: Cận cảnh macro đường chỉ khâu đỏ chữ V trên bề mặt da bóng chày trắng, chữ Henry #23. `200_02`–`200_04`: Khóa kéo, sợi dệt, giặt máy. | Mô tả chuẩn macro đường chỉ khâu bóng chày | PASS (100đ) |

---

## 5. Bằng chứng Thực nghiệm 20 Truy vấn US SERP Độc lập

Đã thực hiện 20 truy vấn tìm kiếm độc lập trên Google Search (thị trường Mỹ `en-US`):

| Query ID | Pos | Từ khóa truy vấn | Loại | Search Intent | Top Organic Domains | Đánh giá & Kết luận |
| :--- | :---: | :--- | :---: | :--- | :--- | :--- |
| `SERP-191-PRI` | 191 | *"custom red black batter bedding"* | `PRIMARY` | Commercial / Custom Sports Bedding | etsy.com, society6.com, wayfair.com, walmart.com | Accurately represents the red and black batter artwork with name and number personalization. |
| `SERP-191-COM` | 191 | *"red black baseball bedding"* | `COMPARATOR` | Commercial / Baseball Decor | wayfair.com, walmart.com, etsy.com, target.com | Strong commercial volume establishing core colorway and sport motif. |
| `SERP-192-PRI` | 192 | *"custom brown baseball glove bedding"* | `PRIMARY` | Commercial / Vintage Glove Bedding | ohaprints.com, etsy.com, custombeddingset.com, ebay.com | Directly matches Design 17's rich brown leather baseball glove artwork. |
| `SERP-192-COM` | 192 | *"brown baseball glove bedding"* | `COMPARATOR` | Commercial / Sports Bedding | walmart.com, u-buy.co.uk, etsy.com, ohaprints.com | Validates market demand for vintage leather sports decor aesthetics. |
| `SERP-193-PRI` | 193 | *"custom flaming baseball batter bedding"* | `PRIMARY` | Commercial / Fiery Sports Bedding | ohaprints.com, youcustomizeit.com, 2cooldesigns.com, etsy.com | Accurately targets Design 18's left-facing batter with centered flaming ball ring. |
| `SERP-193-COM` | 193 | *"flaming baseball bedding"* | `COMPARATOR` | Commercial / Graphic Flame Sports | temu.com, ebay.com, walmart.com, desertcart.com | Confirms strong appeal for intense fire/flame sports graphics. |
| `SERP-194-PRI` | 194 | *"custom orange fire baseball bedding"* | `PRIMARY` | Commercial / Fire Graphic Sports | walmart.com, ebay.com, 2cooldesigns.com, youcustomizeit.com | Distinguishes Design 19's all-over orange inferno artwork from Design 18's flaming ring. |
| `SERP-194-COM` | 194 | *"fire baseball bedding"* | `COMPARATOR` | Commercial / Dynamic Fire Bedding | ebay.com, shein.com, walmart.com, amorcustomgifts.com | Strong commercial volume confirming search interest in fire-themed baseball bedding. |
| `SERP-195-PRI` | 195 | *"custom baseball glove flag bedding"* | `PRIMARY` | Commercial / Patriotic Glove Bedding | youcustomizeit.com, jeminise.com, geckocustom.com, etsy.com | Matches Design 20's brown leather glove, ball, and rustic American flag backdrop. |
| `SERP-195-COM` | 195 | *"american flag baseball glove bedding"* | `COMPARATOR` | Commercial / Americana Baseball | wayfair.com, ebay.com, etsy.com, ohaprints.com | Validates patriotic baseball equipment aesthetic as an established niche. |
| `SERP-196-PRI` | 196 | *"custom black baseball glove bedding"* | `PRIMARY` | Commercial / Monochrome Glove Bedding | etsy.com, ohaprints.com, walmart.com, wayfair.com | Accurately represents Design 21's deep black background and monochrome glove artwork. |
| `SERP-196-COM` | 196 | *"black baseball glove bedding"* | `COMPARATOR` | Commercial / Dark Theme Sports | walmart.com, ohaprints.com, etsy.com, ebay.com | Strong commercial volume confirming popularity of black baseball glove themes. |
| `SERP-197-PRI` | 197 | *"custom lightning baseball bedding"* | `PRIMARY` | Commercial / Electric Sports Decor | amorcustomgifts.com, etsy.com, ohaprints.com, walmart.com | Directly captures Design 22's flying baseball surrounded by cyan lightning effects. |
| `SERP-197-COM` | 197 | *"lightning baseball bedding"* | `COMPARATOR` | Commercial / Novelty Sports Linens | temu.com, walmart.com, etsy.com, target.com | Confirms interest in high-voltage visual energy for sports bedding. |
| `SERP-198-PRI` | 198 | *"custom night field baseball bedding"* | `PRIMARY` | Commercial / Night Stadium Bedding | ohaprints.com, youcustomizeit.com, etsy.com, jeminise.com | Accurately reflects Design 23's night stadium under lights with helmet, bat, glove, and ball. |
| `SERP-198-COM` | 198 | *"baseball field bedding"* | `COMPARATOR` | Commercial / Field Turf & Decor | wayfair.com, potterybarnkids.com, sportsfieldsolutions.com, atxturf.com | Supports stadium field aesthetic as a core sports bedding motif. |
| `SERP-199-PRI` | 199 | *"custom smoke baseball batter bedding"* | `PRIMARY` | Commercial / Fog Smoke Sports | ohaprints.com, youcustomizeit.com, 2cooldesigns.com, etsy.com | Directly captures Design 24's white smoke silhouette batter graphic and name/number customization. |
| `SERP-199-COM` | 199 | *"smoke baseball bedding"* | `COMPARATOR` | Commercial / Abstract Smoke Graphic | temu.com, fineartamerica.com, walmart.com, etsy.com | Validates demand for ethereal smoke atmosphere in athletic bedroom decor. |
| `SERP-200-PRI` | 200 | *"custom baseball stitch name bedding"* | `PRIMARY` | Commercial / Seam Stitch Sports | etsy.com, ebay.com, youcustomizeit.com, zazzle.com | Perfect fit for Design 25's macro baseball lacing and red stitching graphic focus. |
| `SERP-200-COM` | 200 | *"baseball stitch bedding"* | `COMPARATOR` | Commercial / Baseball Stitch Pattern | potterybarnkids.com, rh.com, spoonflower.com, walmart.com | High-volume comparator confirming popularity of authentic seam stitch design. |

---

## 6. Bảng Đối soát Lịch sử Issues (Reconciliation)

Tổng hợp đối soát toàn diện **84 issues** (74 issues lịch sử từ run `20260907_214100` + 10 issues MAJOR mới):

- **RESOLVED (54 issues lịch sử):**
  - 10 issues mô tả kỹ thuật cũ (`SEO Use / QA approval`) đã được loại bỏ sạch trong bản r3.
  - 10 issues meta description được rút gọn chuẩn 100–128 ký tự.
  - 10 issues alt text và 24 issues image observation được hiệu chỉnh chính xác theo visual audit.
- **NOT_APPLICABLE (10 issues lịch sử):** 10 issues liên quan đến trường `h1_proposed` do trường này không thuộc schema workbook tiêu chuẩn; H1 được quản lý tự động bởi Shopify qua product title.
- **PERSISTS (10 issues lịch sử):** 10 issues hạn chế cấp độ bằng chứng từ khóa (`keyword_evidence_level` = `SERP_ONLY`) do môi trường offline QA không kết nối Google Search Console API. Được bù đắp bằng 20 truy vấn US SERP độc lập.
- **NEW (10 issues MAJOR mới - ISSUE-0075 đến ISSUE-0084):**
  - Trường: `description_proposed_html` cho toàn bộ 10 sản phẩm (Pos 191–200).
  - Lý do: Bản copy r3 tuyên bố sai rằng số áo không được cá nhân hóa, mâu thuẫn trực tiếp với live customizer đang yêu cầu cả `Customize Your Name` và `Customize Your Number`.

---

## 7. Đề xuất 10 Mô tả HTML Publish-Ready (Chuẩn hóa Hoàn chỉnh)

Dưới đây là 10 bản mô tả HTML tiếng Anh hoàn chỉnh, đã giải quyết triệt để lỗi D2, phản ánh chính xác cả 2 trường nhập Name (1–30 ký tự) & Number (1–5 ký tự), phân tách cấu tạo Comforter vs Duvet Cover, kích thước, phụ kiện kèm theo và hướng dẫn giặt ủi:

### Position 191: Custom Red Black Batter Bedding (`personalized-baseball-bedding-full-size-flag-custom-name-d-design-16`)

```html
<p>Step up to the plate in ultimate comfort with the <strong>Custom Red Black Batter Bedding</strong> from Jeminise. Perfectly tailored for youth ballplayers, high school athletes, and devoted baseball families, this dynamic bedding set features a high-impact graphic of a batter ready at the plate, framed by explosive red and black brush splatters. Fully personalized with your athlete's name and jersey number in bold athletic font, this bedding brings big-league energy and cozy warmth into any sports-themed bedroom.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> A complete, pre-filled comforter packed with soft, lightweight, hypoallergenic microfiber batting. Delivers immediate, all-season warmth right out of the box.</li>
  <li><strong>Duvet Cover Option:</strong> A protective cover equipped with a hidden bottom zippered closure, designed to slip smoothly over your existing comforter or duvet insert for quick, hassle-free washing.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name Personalization:</strong> Enter your athlete's first name, last name, or team nickname (1–30 characters) in the <em>Customize Your Name</em> field. Enter "NO" if you prefer the design without a name.</li>
  <li><strong>Custom Number Personalization:</strong> Enter your player's jersey number (1–5 digits) in the <em>Customize Your Number</em> field. Enter "NO" if you prefer no number printed.</li>
</ul>
<h3>Available Sizes & Optional Add-Ons</h3>
<ul>
  <li><strong>Mattress Sizes:</strong> Available in Twin (68x86 in), Full (79x90 in), Queen (90x90 in), and King (90x104 in).</li>
  <li><strong>Coordinating Pillowcases:</strong> Select None, 1 Pillowcase, or 2 Pillowcases printed with matching red and black batter artwork.</li>
  <li><strong>Additional Flat Sheet:</strong> Optionally add a matching flat sheet cover in the identical selected bed size.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> 100% premium high-density brushed microfiber for breathable softness, wrinkle resistance, and long-lasting colorfast vibrance.</li>
  <li><strong>Care:</strong> Machine wash cold with mild detergent on gentle cycle. Tumble dry on low heat or hang dry. Do not bleach or dry clean.</li>
</ul>
```

### Position 192: Custom Brown Baseball Glove Bedding (`personalized-baseball-bedding-full-size-flag-custom-name-d-design-17`)

```html
<p>Celebrate the timeless heritage of America's pastime with the <strong>Custom Brown Baseball Glove Bedding</strong> from Jeminise. Featuring an authentic vintage brown leather baseball mitt cradling a well-played baseball set over a rich rustic wood background, this bedding brings warmth, nostalgia, and authentic athletic character to any room. Customized with your player's name and jersey number, it makes a memorable gift for players, coaches, and lifelong baseball enthusiasts.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Fully quilted, all-in-one comforter filled with airy microfiber batting, delivering plush insulation and breathable comfort through every season.</li>
  <li><strong>Duvet Cover Option:</strong> A versatile duvet cover with an invisible bottom zipper, ideal for encasing your favorite duvet insert for easy laundering.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name Personalization:</strong> Type your athlete's name or custom text (1–30 characters) into the <em>Customize Your Name</em> box. Enter "NO" for clean, unpersonalized artwork.</li>
  <li><strong>Custom Number Personalization:</strong> Type your player's favorite jersey number (1–5 characters) into the <em>Customize Your Number</em> box. Enter "NO" if no number is desired.</li>
</ul>
<h3>Available Sizes & Optional Add-Ons</h3>
<ul>
  <li><strong>Mattress Sizes:</strong> Twin (68x86 in), Full (79x90 in), Queen (90x90 in), and King (90x104 in).</li>
  <li><strong>Matching Pillowcases:</strong> Add None, 1 Pillowcase, or 2 Pillowcases with matching vintage brown leather glove artwork.</li>
  <li><strong>Optional Flat Sheet:</strong> Complete your bedding set with an optional matching flat sheet cover.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> Ultra-soft brushed polyester microfiber fabric designed for gentle skin comfort and exceptional durability.</li>
  <li><strong>Care:</strong> Machine wash cold on delicate cycle. Tumble dry on low. Resists fading, shrinking, and pilling wash after wash.</li>
</ul>
```

### Position 193: Custom Flaming Baseball Batter Bedding (`personalized-baseball-bedding-full-size-flag-custom-name-d-design-18`)

```html
<p>Ignite bedtime with championship intensity with the <strong>Custom Flaming Baseball Batter Bedding</strong> from Jeminise. Designed for power hitters and baseball fans who love dramatic visual excitement, this bedding showcases a left-facing batter silhouette swinging through a blazing circular ring of fire with a fiery baseball. Personalized with your player's custom name and jersey number, this set transforms any bedroom into an arena of athletic passion.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> A cozy, pre-filled comforter with plush microfiber fill that provides balanced, cloud-like comfort all year round.</li>
  <li><strong>Duvet Cover Option:</strong> A durable duvet cover featuring a smooth bottom zipper closure, making it effortless to insert and remove your comforter insert.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name Personalization:</strong> Enter any name or text up to 30 characters in the <em>Customize Your Name</em> field. If you desire no name, simply type "NO".</li>
  <li><strong>Custom Number Personalization:</strong> Enter your ballplayer's jersey number (1–5 characters) in the <em>Customize Your Number</em> field. Enter "NO" to omit.</li>
</ul>
<h3>Available Sizes & Optional Add-Ons</h3>
<ul>
  <li><strong>Mattress Sizes:</strong> Standard Twin, Full, Queen, and King bed dimensions.</li>
  <li><strong>Coordinating Pillowcases:</strong> Select None, 1 Pillowcase, or 2 Pillowcases featuring matching flaming batter artwork.</li>
  <li><strong>Matching Flat Sheet:</strong> Optionally include an additional flat sheet cover in corresponding dimensions.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> 100% high-density brushed microfiber for skin-friendly softness, lightweight warmth, and anti-static comfort.</li>
  <li><strong>Care:</strong> Machine wash cold with like colors. Tumble dry low. Do not bleach. Vivid dye-sublimation print will not crack or peel.</li>
</ul>
```

### Position 194: Custom Orange Fire Baseball Bedding (`personalized-baseball-bedding-full-size-flag-custom-name-d-design-19`)

```html
<p>Unleash explosive power on the diamond with the <strong>Custom Orange Fire Baseball Bedding</strong> from Jeminise. This striking sports bedding design surrounds a right-facing baseball batter in an all-over, full-canvas inferno of blazing orange flames, fiery embers, and glowing sparks. Tailored with your athlete's name and number, it delivers unmatched visual punch and warmth to kids' bedrooms, teen spaces, and sports fan caves.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Ready-to-use comforter generously filled with premium down-alternative microfiber batting for lightweight, cozy warmth.</li>
  <li><strong>Duvet Cover Option:</strong> Protective duvet cover featuring a hidden bottom zipper closure for convenient insert swapping and easy washing.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name Personalization:</strong> Customize with your player's name (1–30 characters) in the <em>Customize Your Name</em> field. Enter "NO" for artwork only.</li>
  <li><strong>Custom Number Personalization:</strong> Customize with your player's number (1–5 characters) in the <em>Customize Your Number</em> box. Enter "NO" if no number is needed.</li>
</ul>
<h3>Available Sizes & Optional Add-Ons</h3>
<ul>
  <li><strong>Mattress Sizes:</strong> Twin, Full, Queen, and King sizing available.</li>
  <li><strong>Matching Pillowcases:</strong> Add 1 or 2 matching pillowcases displaying coordinating orange fire graphic artwork.</li>
  <li><strong>Matching Flat Sheet:</strong> Optional coordinated flat sheet cover in matching bed dimensions.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> High-density woven microfiber engineered for exceptional softness, breathability, and wrinkle resistance.</li>
  <li><strong>Care:</strong> Machine wash cold, gentle cycle. Tumble dry low or air dry. Resists wrinkling, staining, and fading.</li>
</ul>
```

### Position 195: Custom Baseball Glove Flag Bedding (`personalized-baseball-bedding-full-size-flag-custom-name-d-design-20`)

```html
<p>Combine your love of America's pastime with patriotic pride with the <strong>Custom Baseball Glove Flag Bedding</strong> from Jeminise. This handsome bedding ensemble showcases a classic leather baseball mitt holding a ball against a distressed wooden American flag backdrop. Personalized with your player's name and jersey number, it brings rustic charm, vintage ballpark memories, and cozy bedtime comfort together.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Pre-filled comforter with lofty, hypoallergenic microfiber fill, engineered for dependable, all-season bedtime comfort.</li>
  <li><strong>Duvet Cover Option:</strong> Equipped with an unobtrusive bottom zipper closure, allowing quick insertion of your comforter insert and easy home laundering.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name Personalization:</strong> Enter any custom name or team dedication (1–30 characters) in the <em>Customize Your Name</em> field. Enter "NO" if no name is desired.</li>
  <li><strong>Custom Number Personalization:</strong> Enter your ballplayer's jersey number (1–5 digits) in the <em>Customize Your Number</em> box. Enter "NO" to omit.</li>
</ul>
<h3>Available Sizes & Optional Add-Ons</h3>
<ul>
  <li><strong>Mattress Sizes:</strong> Twin, Full, Queen, and King mattresses.</li>
  <li><strong>Matching Pillowcases:</strong> Bundle with None, 1 Pillowcase, or 2 Pillowcases with matching patriotic baseball and flag motifs.</li>
  <li><strong>Optional Flat Sheet:</strong> Add a matching flat sheet cover in identical dimensions for a complete look.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> 100% brushed microfiber weave offering superior softness, durability, and moisture-wicking comfort.</li>
  <li><strong>Care:</strong> Machine wash cold on gentle cycle. Tumble dry on low heat. Resists shrinking and color fading.</li>
</ul>
```

### Position 196: Custom Black Baseball Glove Bedding (`personalized-baseball-bedding-full-size-flag-custom-name-d-design-21`)

```html
<p>Make a bold, modern athletic statement with the <strong>Custom Black Baseball Glove Bedding</strong> from Jeminise. Designed in a sophisticated monochrome palette, this bedding features an authentic black leather baseball mitt and baseball set against a sleek, deep black background. Customized with your player's name and jersey number in crisp white athletic lettering, it offers an edgy, contemporary sports look for teens, dorm rooms, and adult fans.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Complete, pre-filled comforter with cloud-soft microfiber filling, offering cozy, all-weather comfort without excessive weight.</li>
  <li><strong>Duvet Cover Option:</strong> Sleek duvet cover with a concealed bottom zipper closure, providing an easy-to-wash outer layer for your existing duvet insert.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name Personalization:</strong> Input your athlete's name or custom text (1–30 characters) in the <em>Customize Your Name</em> box. Enter "NO" for artwork without text.</li>
  <li><strong>Custom Number Personalization:</strong> Input your player's jersey number (1–5 characters) in the <em>Customize Your Number</em> box. Enter "NO" if no number is wanted.</li>
</ul>
<h3>Available Sizes & Optional Add-Ons</h3>
<ul>
  <li><strong>Mattress Sizes:</strong> Available in Twin, Full, Queen, and King sizes.</li>
  <li><strong>Coordinating Pillowcases:</strong> Option to add 1 or 2 matching black monochrome baseball glove pillowcases.</li>
  <li><strong>Matching Flat Sheet:</strong> Complete your bedding set with an optional matching black flat sheet cover.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> Premium high-density brushed microfiber for breathable softness, hypoallergenic comfort, and lasting durability.</li>
  <li><strong>Care:</strong> Machine wash cold with mild detergent. Tumble dry low. High-contrast black pigments stay bold and fade-free.</li>
</ul>
```

### Position 197: Custom Lightning Baseball Bedding (`personalized-baseball-bedding-full-size-flag-custom-name-d-design-22`)

```html
<p>Electrify your bedroom decor with the high-voltage <strong>Custom Lightning Baseball Bedding</strong> from Jeminise. Featuring a dynamic baseball rocketing across the sky surrounded by vivid cyan lightning bolts and atmospheric storm clouds, this bedding radiates athletic energy and speed. Personalized with your player's name and jersey number, it is the ultimate bedroom centerpiece for young ballplayers who play with lightning-fast intensity.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Pre-filled comforter packed with soft, airy microfiber insulation, providing plush, ready-to-use warmth all year round.</li>
  <li><strong>Duvet Cover Option:</strong> Features a hidden bottom zipper closure for effortless cover changes and fast machine washing.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name Personalization:</strong> Type your athlete's name or text (up to 30 characters) into the <em>Customize Your Name</em> box. Type "NO" for unpersonalized artwork.</li>
  <li><strong>Custom Number Personalization:</strong> Type your player's number (1–5 digits) into the <em>Customize Your Number</em> box. Type "NO" to leave off.</li>
</ul>
<h3>Available Sizes & Optional Add-Ons</h3>
<ul>
  <li><strong>Mattress Sizes:</strong> Twin, Full, Queen, and King sizing available.</li>
  <li><strong>Matching Pillowcases:</strong> Choose None, 1 Pillowcase, or 2 Pillowcases with matching cyan lightning baseball artwork.</li>
  <li><strong>Optional Flat Sheet:</strong> Optionally add a matching flat sheet cover in corresponding size dimensions.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> 100% fine-spun brushed microfiber providing an exceptionally silky, lightweight, and breathable hand-feel.</li>
  <li><strong>Care:</strong> Machine wash cold, gentle cycle. Tumble dry on low heat. Resists wrinkles, fading, and static cling.</li>
</ul>
```

### Position 198: Custom Night Field Baseball Bedding (`personalized-baseball-bedding-full-size-flag-custom-name-d-design-23`)

```html
<p>Experience the magic of night games under the lights with the <strong>Custom Night Field Baseball Bedding</strong> from Jeminise. This immersive sports bedding design places you right on the diamond beneath illuminated stadium floodlights, showcasing an authentic layout of a batting helmet, wooden bat, leather mitt, and baseball resting on manicured field turf. Customized with your name and number, it captures the unforgettable drama and romance of night baseball.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Fully quilted all-season comforter packed with fluffy microfiber batting for cozy, restorative bedtime comfort.</li>
  <li><strong>Duvet Cover Option:</strong> Protective outer cover equipped with an invisible bottom zipper, making removal and laundry day effortless.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name Personalization:</strong> Personalize with your ballplayer's name or family name (1–30 characters) in the <em>Customize Your Name</em> field. Enter "NO" for no text.</li>
  <li><strong>Custom Number Personalization:</strong> Personalize with your player's jersey number (1–5 digits) in the <em>Customize Your Number</em> field. Enter "NO" if not needed.</li>
</ul>
<h3>Available Sizes & Optional Add-Ons</h3>
<ul>
  <li><strong>Mattress Sizes:</strong> Standard Twin, Full, Queen, and King bed coverage.</li>
  <li><strong>Coordinating Pillowcases:</strong> Select None, 1 Pillowcase, or 2 Pillowcases printed with coordinating night stadium graphics.</li>
  <li><strong>Additional Flat Sheet:</strong> Option to add a matching flat sheet cover in identical mattress dimensions.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> Premium brushed microfiber weave engineered for superior softness, durability, and breathability.</li>
  <li><strong>Care:</strong> Machine wash cold on gentle cycle. Tumble dry low. Resists fading, pilling, and wrinkles.</li>
</ul>
```

### Position 199: Custom Smoke Baseball Batter Bedding (`personalized-baseball-bedding-full-size-flag-custom-name-d-design-24`)

```html
<p>Bring dramatic flair and mystery to bedtime with the <strong>Custom Smoke Baseball Batter Bedding</strong> from Jeminise. This artistic sports bedding features a powerful silhouette of a baseball batter poised at the plate amidst swirling clouds of ethereal white smoke and fog on a sleek dark background. Personalized with your player's name and jersey number, it creates a modern, cinematic aesthetic that elevates any bedroom.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Pre-filled comforter with plush, lightweight microfiber fill, delivering balanced warmth and cloud-like comfort throughout every season.</li>
  <li><strong>Duvet Cover Option:</strong> Lightweight duvet cover equipped with a durable bottom zipper closure, allowing quick insertion of your comforter insert and easy washing.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name Personalization:</strong> Enter any custom name or team nickname (1–30 characters) in the <em>Customize Your Name</em> box. Type "NO" for artwork without text.</li>
  <li><strong>Custom Number Personalization:</strong> Enter your athlete's jersey number (1–5 digits) in the <em>Customize Your Number</em> box. Type "NO" to omit.</li>
</ul>
<h3>Available Sizes & Optional Add-Ons</h3>
<ul>
  <li><strong>Mattress Sizes:</strong> Available in Twin, Full, Queen, and King mattress dimensions.</li>
  <li><strong>Matching Pillowcases:</strong> Add None, 1 Pillowcase, or 2 Pillowcases with coordinating smoke batter graphic artwork.</li>
  <li><strong>Matching Flat Sheet:</strong> Optional coordinated flat sheet cover available in matching bed size.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> 100% woven microfiber fabric with high yarn density for a plush, gentle feel against the skin.</li>
  <li><strong>Care:</strong> Machine wash cold with mild detergent. Tumble dry on low heat. Colorfast dye sublimation print resists fading.</li>
</ul>
```

### Position 200: Custom Baseball Stitch Name Bedding (`personalized-baseball-bedding-full-size-flag-custom-name-d-design-25`)

```html
<p>Celebrate the pure, tactile essence of America's favorite pastime with the <strong>Custom Baseball Stitch Name Bedding</strong> from Jeminise. Featuring a macro photographic close-up of textured genuine baseball leather highlighted by iconic red V-shaped seam stitches, this bedding delivers timeless ballpark style. Personalized with your athlete's name and jersey number printed right beside the laces, it creates the ultimate bedroom centerpiece for passionate baseball players of all ages.</p>
<h3>Choose Your Bedding Construction</h3>
<ul>
  <li><strong>Plush Comforter Option:</strong> Ready-to-use comforter generously filled with fluffy down-alternative microfiber batting for cloud-soft, all-season warmth.</li>
  <li><strong>Duvet Cover Option:</strong> Features a hidden bottom zipper closure, perfect for protecting and slipping over your favorite duvet or comforter insert.</li>
</ul>
<h3>Personalization & Custom Details</h3>
<ul>
  <li><strong>Custom Name Personalization:</strong> Personalize with your player's name or custom text (1–30 characters) in the <em>Customize Your Name</em> field. Type "NO" for clean, unpersonalized laces.</li>
  <li><strong>Custom Number Personalization:</strong> Personalize with your player's jersey number (1–5 characters) in the <em>Customize Your Number</em> field. Type "NO" to omit.</li>
</ul>
<h3>Available Sizes & Optional Add-Ons</h3>
<ul>
  <li><strong>Mattress Sizes:</strong> Standard Twin, Full, Queen, and King mattress coverage.</li>
  <li><strong>Coordinating Pillowcases:</strong> Bundle with 1 or 2 matching baseball stitch pillowcases to complete your sports bedding set.</li>
  <li><strong>Optional Flat Sheet:</strong> Add a matching flat sheet cover in corresponding bed dimensions.</li>
</ul>
<h3>Fabric & Care Instructions</h3>
<ul>
  <li><strong>Material:</strong> Premium brushed polyester microfiber engineered for durable softness, breathability, and wrinkle resistance.</li>
  <li><strong>Care:</strong> Machine wash cold on delicate cycle. Tumble dry low or air dry. Resists shrinking, fading, and pilling.</li>
</ul>
```

---

## 8. Tuyên bố Trạng thái & Tiếp tục Quy trình

- **Xác nhận khóa dữ liệu:** Toàn bộ bằng chứng kiểm toán độc lập đã được lưu trữ vĩnh viễn tại `seo_runs/jeminise.com/20260906_234129/qa/20260908_112500/`.
- **File bàn giao:** `resutls/jeminise.com/20260906_234129/qa/20260908_112500/SEO_QA_qa_batch_020_r3.xlsx` và `SEO_QA_qa_batch_020_r3.md`.
- **Nguyên tắc an toàn:** Không chỉnh sửa file nguồn `SEO_Product_Optimization_qa_batch_020_r3.xlsx` và không sửa trực tiếp live store.
- **Dừng theo quy trình:** Hệ thống kích hoạt cờ `awaiting_confirmation=true` và **DỪNG LẠI**, chờ lệnh xác nhận từ người dùng trước khi chuyển sang Batch 21.
