# SOP triển khai Shopify On-Page từ workbook SEO

Phiên bản 1.1 — bổ sung QA nội dung có chấm điểm ngày **2026-09-06**; mapping kỹ thuật đối chiếu tài liệu chính thức ngày **2026-09-05**.

Đầu vào: workbook tạo theo [prompt.md](./prompt.md), được kiểm tra bằng [prompt_qa.md](./prompt_qa.md) trước khi duyệt. Đối tượng sử dụng: người duyệt nội dung, người vận hành Shopify và kỹ thuật viên triển khai. Giao diện minh họa dùng nhãn tiếng Anh; shop dùng ngôn ngữ khác cần đối chiếu chức năng tương ứng.

## 1. Phạm vi và tiêu chí nghiệm thu

Quy trình này cập nhật SEO title, meta description, Product Title dùng cho H1, mô tả sản phẩm và alt text ảnh đã được duyệt. Không thay đổi giá, SKU, tồn kho, tùy chọn biến thể, shipping, fulfillment, thuế hoặc tình trạng xuất bản. Product Title và Description là nội dung sản phẩm được dùng ở nhiều nơi; chúng không phải các trường chỉ có tác động đến Google.

Mục tiêu “Zero-Downtime / No-Risk” được thực thi bằng giới hạn trường, backup, kiểm tra khóa, triển khai thử và rollback. Không có cơ sở cam kết rủi ro bằng 0 hay Google giữ nguyên thứ hạng. Nghiệm thu kỹ thuật yêu cầu **100% trường đã ghi khớp phiên bản được duyệt**, không có sản phẩm/ảnh mới ngoài kế hoạch và không có thay đổi ngoài phạm vi do job gây ra. Nếu chưa có bằng chứng đối chiếu thì chưa nghiệm thu.

Không cần đóng shop, đổi theme hoặc bật password để sửa nội dung. Nếu có app/Flow tự phản ứng với thay đổi sản phẩm, kiểm tra phạm vi tác động trước; không tự tắt app hoặc luồng bán hàng. Việc cập nhật theme/metafield đặc thù để sửa H1 nằm ngoài file import tiêu chuẩn, phải ghi thành thay đổi riêng.

### Các khái niệm phải phân biệt

| Khái niệm | Cách hiểu khi triển khai |
|---|---|
| Product Title | Tên sản phẩm trong Shopify; theme thường dùng để render H1, collection card và các vị trí khác |
| SEO title | Giá trị Page title trong Search engine listing; theme dùng để tạo thẻ HTML `<title>` |
| Google title link | Tiêu đề thực tế Google hiển thị; có thể được Google tạo lại từ nhiều nguồn |
| Meta description | Giá trị tạo thẻ `<meta name="description">`; có thể được Google dùng làm snippet |
| Product Description | Nội dung HTML trên trang sản phẩm; không phải meta description |
| Alt text | Mô tả media được lưu trong Shopify và cần được theme đưa vào thuộc tính `alt` của `<img>` |
| Handle | Slug hiện tại, ví dụ `floral-throw-blanket`; không phải cả URL |
| Product ID | ID của sản phẩm trong đúng shop; khác Variant ID, Media ID và GraphQL GID |

Google có thể tạo lại title link và đoạn trích. SERP Preview chỉ mô phỏng; không chứng minh kết quả đã cập nhật trên Google. [Google: title links](https://developers.google.com/search/docs/appearance/title-link), [Google: snippets](https://developers.google.com/search/docs/appearance/snippet).

## 2. Vai trò, duyệt và giới hạn lô

1. **Agent nghiên cứu:** tạo đề xuất và bằng chứng; mặc định `NEEDS_REVIEW`, không tự duyệt.
2. **Agent QA:** dùng prompt_qa.md để đọc lại nguồn và từng ảnh, chấm bản workbook xác định, lập lỗi và đề nghị sửa. Không sửa workbook nguồn hoặc tự duyệt.
3. **Người duyệt:** đọc báo cáo QA, quyết định từng trường, ghi tên, thời điểm và revision.
4. **Người triển khai:** tạo file chỉ chứa thay đổi đã duyệt, backup, nhập và QA sau triển khai. Người duyệt và người triển khai có thể là cùng người, nhưng vẫn ghi nhật ký.

`review_status` trong `SEO_Products` chỉ nhận:

| Giá trị | Xử lý |
|---|---|
| `APPROVED` | Được xét để xuất payload; còn phải qua kiểm tra khóa, revision và giá trị gốc |
| `NEEDS_REVIEW` | Không đưa vào file triển khai |
| `KEEP_ORIGINAL` | Giữ sản phẩm hiện tại; không đưa vào file triển khai |

`QA_PASSED`, `PARTIAL`, `BLOCKED` thuộc `processing_status`, không thay cho quyết định duyệt. Keyword chưa có dữ liệu ngoài vẫn có thể được người duyệt chấp nhận như giả thuyết; phải giữ nhãn mức kiểm chứng thật.

Ví dụ một dòng được duyệt:

```text
Handle: floral-throw-blanket
revision: r3
review_status: APPROVED
approved_by: <tên người thực sự duyệt>
approved_at: <thời điểm ISO 8601 có múi giờ>
approved_revision: r3
approved_fields: meta_title_seo,meta_description_seo
title_action: KEEP
meta_title_action: SET
meta_description_action: SET
description_action: KEEP
```

Chỉ hai trường meta được ghi. Một ô `title_proposed` có nội dung không tự cấp quyền đổi Product Title. Với ảnh, phải duyệt cả dòng sản phẩm và dòng `Image_Audit`, cùng đúng revision và trường `img_N_alt`/`alt_proposed` tương ứng.

Agent nghiên cứu tuân thủ **10 sản phẩm/lô, tối đa 10 sản phẩm/lượt** theo prompt.md; lô cuối có thể ít hơn 10. Lưu sheet trong `resutls/<shop-domain>/<run_id>/` theo quy ước của prompt, lưu tiến độ rồi chờ xác nhận. Xác nhận lô nghiên cứu không phải APPROVED hay lệnh import. Trong SOP này, mỗi job triển khai tối đa 20 sản phẩm; nhóm thử đầu tiên 3–5 sản phẩm nằm trong số đã duyệt. Catalog hàng nghìn sản phẩm được chia nhiều job có manifest, không gộp thành một lần ghi không kiểm soát.

Luồng 1 áp dụng dưới 20 sản phẩm; với đúng 20 có thể dùng cùng luồng. Luồng 2 áp dụng khi tổng danh sách trên 20 sản phẩm và chia thành các job tối đa 20. Con số này là quy định vận hành của dự án, không phải giới hạn Shopify/Matrixify.

### 2.1. QA chất lượng nội dung trước khi duyệt

Sau mỗi lô nghiên cứu, chạy [prompt QA](./prompt_qa.md) trên đúng 10 sản phẩm của lô đó (lô cuối có thể ít hơn). Workbook nghiên cứu tích lũy có thể chứa nhiều lô; truyền danh sách product_key hoặc batch_id để QA không chấm lại/toàn bộ ngoài ý định. QA là lượt riêng, chỉ bắt đầu khi được người dùng yêu cầu; sau lô QA lưu kết quả và chờ xác nhận lô tiếp theo.

Ví dụ lệnh dùng, thay các placeholder bằng đường dẫn và mã thật:

```text
Hãy thực hiện prompt_qa.md.
Workbook nguồn: resutls/<shop-domain>/<run_id>/batches/SEO_Product_Optimization_through_<batch_id>.xlsx
Hồ sơ nghiên cứu: seo_runs/<shop-domain>/<run_id>/
Phạm vi QA: đúng các product_key của <batch_id>, tối đa 10 sản phẩm.
Đối chiếu lại trang và từng ảnh, chấm điểm, lưu báo cáo; không sửa workbook nguồn hoặc Shopify.
```

QA chấm 100 điểm/sản phẩm: đúng sản phẩm/thuộc tính 25; keyword/intent/bằng chứng nhu cầu 20; SEO title và H1 15; meta và mô tả 15; ảnh/alt 20; bằng chứng và nhất quán dữ liệu 5. Từng ảnh có điểm riêng 100. Công thức chi tiết, xử lý phần chưa kiểm tra và mức lỗi nằm trong prompt_qa.md, là nguồn chuẩn duy nhất cho rubric.

| Kết luận QA | Ý nghĩa và bước tiếp theo |
|---|---|
| QA_PASS | Đủ độ phủ, điểm >=85, không CRITICAL/MAJOR; chuyển người duyệt, chưa phải APPROVED |
| QA_REVISE | Đủ dữ liệu, điểm 70–<85 hoặc còn MAJOR, không thuộc QA_FAIL; sửa và QA lại |
| QA_FAIL | Có CRITICAL đã xác định hoặc điểm cuối <70; trả về sửa, không dùng điểm trung bình lô để bỏ qua |
| QA_INCOMPLETE | Chưa đủ dữ liệu để chấm cuối, không có CRITICAL đã xác định; bổ sung nguồn/công cụ, báo khoảng điểm và phần còn thiếu |

Không tự chuyển QA_PASS thành review_status/content_qa_status trong workbook nguồn. Báo cáo QA và bản nghiên cứu có revision/hash riêng; khi sửa đề xuất, cập nhật revision theo prompt.md, làm mất hiệu lực duyệt cũ và QA lại bản mới. Không sửa báo cáo cũ để làm đẹp điểm. Điểm QA không phải dự báo rank hoặc thay thế việc kiểm tra dữ liệu thực tế sau import ở mục 8.

Với QA_REVISE/QA_FAIL, xử lý lỗi ảnh hưởng các trường dự định triển khai trước khi duyệt. Nếu chỉ thiếu dữ liệu keyword ngoài shop, người duyệt có thể chấp nhận giới hạn một cách rõ ràng, ghi trường/revision và lý do trong nhật ký; báo cáo vẫn giữ QA_INCOMPLETE/mức bằng chứng thật. Ngoại lệ này không cho phép bỏ qua việc xác định sản phẩm/ảnh, thuộc tính hoặc dữ liệu gốc cần cho payload.

Báo cáo nằm tại `resutls/<shop-domain>/<run_id>/qa/<qa_run_id>/SEO_QA_<qa_batch_id>.xlsx` và bản tóm tắt `.md`. Bằng chứng, manifest và qa_progress nằm trong `seo_runs/<shop-domain>/<run_id>/qa/<qa_run_id>/`. Workbook SEO cuối vẫn mang tên `SEO_Product_Optimization.xlsx`; không trộn báo cáo QA vào file import.

## 3. Backup bắt buộc và chuẩn bị dữ liệu gốc

### 3.1. Export Shopify CSV trước mọi đợt chỉnh sửa

1. Đăng nhập đúng Shopify store; ghi domain `.myshopify.com`, thị trường và ngôn ngữ mặc định.
2. Vào **Products → Export**.
3. Chọn **All products** cho backup đầu đợt, hoặc tập sản phẩm được lọc nếu đang làm backup bổ sung của job. Nếu export lớn bị timeout, xuất các tập không chồng lặp rồi đối chiếu độ phủ.
4. Chọn định dạng CSV cho Excel/Numbers/spreadsheet hoặc Plain CSV phù hợp công cụ xử lý.
5. Bấm **Export products**, tải file hoặc lấy file từ email Shopify khi export xong.
6. Lưu bản nguyên gốc dưới `backups/<shop>/<run_id>/products_before_<timestamp>.csv`. Không chỉnh sửa, không sort bản backup.
7. Kiểm tra mở được file và đủ Handle/variant/ảnh trong phạm vi. Tạo bản làm việc riêng.

Đây là các chức năng export được Shopify hướng dẫn. CSV có thể để trống SEO fields khi đang dùng fallback và không chứa bản nhị phân của ảnh. [Shopify: Exporting products](https://help.shopify.com/en/manual/products/import-export/export-products).

### 3.2. Lưu thêm những dữ liệu CSV không chứng minh đủ

- Nếu dùng Matrixify, export Products với Basic Columns, Media và SEO Metafields để lấy ID, URL media gốc, alt, Title, Body HTML, SEO fields.
- Lưu nội dung HTML/render trước sửa: H1, `<title>`, meta description, canonical và ảnh/alt liên quan.
- Giữ bằng chứng định danh variant/media nếu cần QA ID; CSV gốc không phải backup hoàn chỉnh mọi tài nguyên/ID của shop.
- Với SEO field trống, lưu `EMPTY` và nội dung render tương ứng; không thay trống bằng nội dung render rồi gọi đó là giá trị admin gốc.
- Với trường chưa đọc được, ghi `UNKNOWN` và chưa cho xuất payload cho trường đó đến khi đối chiếu được.

Backup đầy đủ không được dùng làm file re-import thường lệ: nó có thể chứa giá, tồn kho hoặc dữ liệu đã cũ. Rollback tại mục 10 chỉ lấy lại các trường đã thay đổi.

### 3.3. Đối chiếu trước khi ghi

Ngay trước mỗi job, lấy export/giá trị admin mới cho các sản phẩm trong job. Kiểm tra:

- Mỗi khóa workbook khớp đúng một sản phẩm hiện có trong đúng shop.
- Mỗi trường SET vẫn bằng giá trị gốc đã dùng để duyệt. Nếu có người/app đã sửa sau đó, chuyển `NEEDS_REVIEW`; không ghi đè bản mới.
- ID và Handle nếu cùng có phải chỉ cùng sản phẩm. ID luôn lưu như text, không dùng dạng `1.234E+12` hoặc số bị làm tròn.
- `revision == approved_revision`; approved_fields chỉ gồm các trường thật sự được phép sửa.
- Số ảnh, URL gốc, biến thể và nơi ảnh xuất hiện vẫn đúng.
- Không có kế hoạch sửa đồng thời cùng sản phẩm bằng người/app/job khác. Không cần ngừng checkout; tồn kho biến động do đơn hàng vẫn là hoạt động bình thường.

CSV/import thông thường không phải giao dịch khóa toàn bộ dữ liệu theo kiểu compare-and-swap. Đối chiếu sát lúc ghi và giảm ghi các trường KEEP là biện pháp hạn chế xung đột; phát hiện thay đổi thì dừng job liên quan.

## 4. Mapping dữ liệu Excel → Shopify → CSV/Matrixify

Shopify đang có tên header mới và duy trì tương thích tên cũ. **Export mới từ chính shop là mẫu chuẩn**; không trộn hai alias của cùng trường trong một file. Matrixify dùng template riêng. Các tên trong bảng dưới đã đối chiếu với [Shopify CSV](https://help.shopify.com/en/manual/products/import-export/using-csv), [Matrixify Products](https://matrixify.app/documentation/products/) và [Matrixify SEO fields](https://matrixify.app/tutorials/how-to-bulk-update-shopify-seo-title-and-description/).

| Cột workbook | Shopify Admin/đích thực tế | Shopify CSV: hiện tại; tên cũ nếu khác | Matrixify: sheet `Products` | Mục đích |
|---|---|---|---|---|
| `Handle` | Search engine listing → URL handle; chỉ dùng định danh | `URL handle`; cũ `Handle` | `Handle` nếu chọn nhận diện bằng Handle | Khóa sản phẩm, giữ URL |
| `product_id` | Product ID đã đối chiếu | Không dùng Product ID làm khóa ở importer CSV gốc | `ID` | Nhận diện sản phẩm; không đổi Handle |
| `title_proposed` | Ô **Title** ở đầu trang sản phẩm; kiểm tra nguồn H1 của theme | `Title` | `Title` | Tên sản phẩm, On-page H1 |
| `meta_title_seo` | Search engine listing → **Page title** | `SEO title`; cũ `SEO Title` | `Metafield: title_tag [string]` | Mô tả trang trong `<title>`, hỗ trợ mức liên quan/CTR SERP |
| `meta_description_seo` | Search engine listing → **Meta description** | `SEO description`; cũ `SEO Description` | `Metafield: description_tag [string]` | Tóm tắt có thể được dùng làm snippet, hỗ trợ CTR |
| `description_proposed_html` | **Description**, toàn bộ HTML cuối cùng | `Description`; cũ `Body (HTML)` | `Body HTML` | Nội dung On-page, thông số và nhu cầu mua |
| `img_N_link` / `image_url_export` | Media hiện có đã được xác minh | `Product image URL`; cũ `Image Src` | `Image Src` | Định vị ảnh, không tải ảnh mới trong đợt alt-only |
| `img_N_alt` / `alt_proposed` | Media → Preview → Add/Edit alt text | `Image alt text`; cũ `Image Alt Text` | `Image Alt Text` | Khả năng tiếp cận và Google Image Search |
| `image_number` | Thứ tự gallery, chỉ đối chiếu | `Image position`; cũ `Image Position` nếu file ảnh cần giữ thứ tự | `Image Position` chỉ khi chủ ý duyệt thay vị trí; SOP alt-only không gửi | Giữ thứ tự ảnh |
| `media_id` | ID media/file dùng đối chiếu | Không tự thêm cột ID ảnh vào CSV gốc | Không tự suy ra cột import Image ID; template SOP dùng Image Src | QA, không phải payload tự động |
| `meta_keyword`, `primary_keyword`, `secondary_keywords` | Không có trường cần nhập cho Google SEO | Không xuất | Không xuất | Nghiên cứu nội bộ, không map sang Tags |
| `review_status`, action, approval, evidence, revision | Quản trị workbook | Không xuất | Không xuất | Quyết định lọc và nhật ký |
| Các cột `*_current`, `rendered_*`, `theme_title_suffix` | Dữ liệu đối chiếu | Không ghi giá trị render vào trường admin theo suy đoán | Không xuất trực tiếp | Baseline và rollback |

Lưu ý về H1: thay Product Title thường thay H1 nhưng cũng có thể thay tên trong collection/feed. Nếu theme dùng metafield/app để tạo H1, ghi mapping riêng và chuyển phần H1 sang NEEDS_REVIEW; không nhập nhầm H1 vào Page title. Theme cũng có thể thêm tên shop vào `<title>`. [Shopify: SEO metadata trong theme](https://shopify.dev/docs/storefronts/themes/seo/metadata).

Với ảnh trong HTML mô tả (`image_location=DESCRIPTION`), alt có thể nằm trực tiếp trong `<img alt="...">` của Description. Cần sửa đúng HTML được duyệt; không giả định đổi alt gallery sẽ đổi alt trong Description. Nếu ảnh là tài nguyên dùng chung nhiều sản phẩm, kiểm tra ảnh hưởng và thống nhất alt trước khi ghi.

### 4.1. Quy tắc giá trị và danh sách trường cho phép

- `KEEP`: không gửi trường đó. `SET`: gửi đúng giá trị đã duyệt, không rỗng.
- Không đưa chữ `Keep`, `UNKNOWN`, `N/A`, chú thích hay công thức Excel vào nội dung shop.
- Không dùng ô rỗng để biểu đạt KEEP. Cột tùy chọn có mặt nhưng rỗng có thể xóa giá trị hiện tại. Bỏ hẳn cột mới biểu đạt không cập nhật khi importer hỗ trợ. [Shopify: Overwriting CSV values](https://help.shopify.com/en/manual/products/import-export/import-products).
- Chia file theo tập trường SET: file chỉ đổi meta title khác file đổi cả title và description. Với Matrixify, điều này tránh gửi các cột KEEP; với native CSV vẫn phải có Title và cấu trúc bắt buộc được giữ nguyên từ bản mới.
- Payload không có Price, Compare-at price, SKU, Barcode, inventory, weight, fulfillment, shipping, tax, Status, Published, Included/market, Tags, Vendor, Product category, Template hoặc custom metafield không được duyệt.
- Ngoại lệ cấu trúc native CSV: Title, các cột option/variant bắt buộc để giữ cấu trúc phải được sao chép nguyên trạng; không coi chúng là nội dung được phép sáng tác.
- Giữ URL Handle và CDN URL. Không đổi filename, thêm ảnh hoặc reorder ảnh trong tác vụ alt-only.
- Chỉ sửa ngôn ngữ mặc định bằng product CSV/template này. Bản dịch theo Markets/Translate & Adapt cần luồng dịch đã xác minh riêng.

Giới hạn trường Shopify theo tài liệu hiện tại: SEO title tối đa 70 ký tự, SEO description tối đa 320, alt tối đa 512. Mốc biên tập đề xuất lần lượt khoảng 50–60, 140–160 và thường không quá 125 ký tự nếu đủ ý. Đếm cả khoảng trắng/dấu câu; không coi các mốc biên tập là giới hạn hiển thị cố định của Google. Kiểm tra lại giới hạn ở export/UI trước một đợt mới. [Shopify CSV fields](https://help.shopify.com/en/manual/products/import-export/using-csv), [Shopify alt text](https://help.shopify.com/en/manual/products/product-media/add-alt-text).

### 4.2. Khác biệt định dạng

Workbook duyệt có sheet `SEO_Products` và bằng chứng. Không upload trực tiếp vào Shopify. Tạo file triển khai tách riêng:

```text
deploy/<job_id>/native_meta.csv                 # Native Shopify CSV
deploy/<job_id>/matrixify_meta.xlsx             # Chỉ sheet Products
deploy/<job_id>/matrixify_image_alt.xlsx        # Chỉ sheet Products
deploy/<job_id>/manifest.json                  # Khóa + trường + giá trị + revision
deploy/<job_id>/rollback_source.json           # Giá trị gốc, kể cả EMPTY
```

Không tạo file triển khai nếu không có dòng hợp lệ APPROVED. UTF-8 và dấu phẩy cho CSV; dùng thư viện CSV để quote đúng dấu phẩy, dấu nháy và xuống dòng. Giữ ID dạng TEXT. Workbook không thực thi chuỗi bắt đầu bằng `=` như công thức; xuất payload từ raw value, không thêm ký tự bảo vệ vào nội dung thật.

## 5. Luồng 1 — sửa thủ công tối đa 20 sản phẩm

### 5.1. Mở đúng sản phẩm

1. Filter workbook `review_status=APPROVED` và kiểm tra approved_revision/approved_fields.
2. Trong Shopify Admin vào **Products**, tìm tên/Handle, mở sản phẩm.
3. Đối chiếu Handle ở Search engine listing và Product ID nếu có. Không chọn chỉ vì tên giống nhau.
4. So sánh giá trị hiện tại với bản đã duyệt. Khác biệt thì dừng sản phẩm đó để xem lại.

Ví dụ xuyên suốt SOP: Handle giả định `floral-throw-blanket`; đây không phải sản phẩm đã được kiểm tra của shop. Mọi ID, URL ảnh và nội dung mẫu phải thay bằng dữ liệu thật đã được duyệt.

### 5.2. SEO title và meta description

1. Cuộn xuống **Search engine listing** ở cuối trang sản phẩm.
2. Bấm biểu tượng chỉnh sửa hoặc **Edit website SEO** tùy giao diện.
3. Điền `meta_title_seo` vào **Page title** nếu trường này được duyệt SET.
4. Điền `meta_description_seo` vào **Meta description** (một số giao diện ghi Description trong khối này).
5. Không chỉnh **URL handle**.
6. Kiểm tra preview, rồi bấm **Save**. Mở lại để xác nhận giá trị đã lưu.

Ví dụ Page title: `Floral Throw Blanket with Blue Botanical Pattern`. Meta description phải mô tả đúng chất liệu/công dụng có nguồn; không chèn free shipping hoặc machine washable chỉ để đủ ký tự. Thao tác SEO cấp sản phẩm được minh họa trong [hướng dẫn Matrixify kiểm tra Shopify Admin](https://matrixify.app/tutorials/how-to-bulk-update-shopify-seo-title-and-description/); quy tắc title/description theo [Shopify SEO fields](https://help.shopify.com/en/manual/promoting-marketing/seo/adding-keywords).

### 5.3. Product Title/H1 và Description

1. Chỉ khi `title_proposed` được duyệt: sửa ô **Title** ở đầu trang thành giá trị đề xuất, ví dụ `Blue Floral Throw Blanket`.
2. Nếu Description được duyệt, dùng editor nội dung sản phẩm; chọn chế độ HTML khi dán `description_proposed_html` và kiểm tra lại chế độ trực quan.
3. APPEND/REPLACE_SECTION phải đã được dựng thành toàn bộ HTML cuối cùng trước khi duyệt. Không dán đoạn bổ sung đè toàn bộ mô tả.
4. Giữ bảng size, thông số, hướng dẫn dùng và liên kết cần thiết theo bản duyệt. Bấm **Save**.
5. Mở storefront: H1 phải lấy đúng nội dung dự kiến, bố cục/định dạng không vỡ. Nếu không đúng do theme, ghi lỗi mapping, không tiếp tục sửa các sản phẩm cùng template.

### 5.4. Alt text từng ảnh

1. Trong sản phẩm, tìm khối **Media**; mở đúng thumbnail sau khi đối chiếu ảnh và URL với `Image_Audit`.
2. Tại Preview media, chọn **Add alt text** hoặc chỉnh alt hiện có.
3. Dán alt đã duyệt, bấm **Save alt text**, đóng preview bằng **X**.
4. Lặp lại với từng ảnh được duyệt. Không chỉ sửa ảnh đầu tiên rồi đánh dấu cả gallery hoàn tất.
5. Mở lại media để kiểm tra giá trị đã lưu, rồi kiểm tra HTML storefront theo mục 8.

Ví dụ alt ảnh tổng thể: `Blue botanical floral throw blanket draped over a cream sofa`. Chỉ dùng nếu ảnh thật thể hiện đúng các chi tiết đó. Alt của ảnh cận cảnh phải mô tả góc cận cảnh, không chép câu này cho mọi ảnh. [Shopify: Add alt text to media](https://help.shopify.com/en/manual/products/product-media/add-alt-text).

## 6. Luồng 2A — hàng loạt bằng Shopify CSV gốc

### 6.1. Khóa, giới hạn và chuẩn bị

CSV gốc cập nhật sản phẩm bằng **Handle** với tùy chọn **Overwrite products with matching handles**. Không tự thêm Product ID rồi kỳ vọng importer dùng ID. Handle không match có thể trở thành sản phẩm mới; phải đối chiếu tập khóa với export hiện tại trước khi upload.

Không có chế độ UPDATE-only tương đương Matrixify trong luồng này. File native cần `URL handle`/`Handle` và `Title`; cấu trúc variant cần giữ đúng option. Tài liệu Shopify cảnh báo thiếu hoặc thay option có thể làm thay đổi/xóa variant. Vì vậy, SOP yêu cầu giữ nguyên toàn bộ cấu trúc option cần thiết từ export, kể cả sản phẩm chỉ có Default Title. [Shopify CSV schema](https://help.shopify.com/en/manual/products/import-export/using-csv).

### 6.2. Tạo file text/meta

1. Đọc bản export mới bằng công cụ không làm mất kiểu dữ liệu hoặc quan hệ các dòng.
2. Lấy toàn bộ block dòng của từng Handle đã duyệt. Không sort riêng một cột; không mất liên kết dòng ảnh/variant với sản phẩm.
3. Giữ Handle, Title và các cột Option1/2/3 name/value cùng dữ liệu liên kết option bắt buộc nếu export có. Giữ chính xác số tổ hợp, giá trị và cách biểu diễn continuation rows của export.
4. Loại các cột giá, SKU, inventory, shipping và các trường ngoài phạm vi theo mục 4.1. Không để chúng tồn tại với ô rỗng.
5. Chỉ thêm các cột text/meta được duyệt SET. Nếu Title không đổi, dùng Title quản trị mới nhất, không lấy H1 render.
6. Điền trường cấp sản phẩm ở dòng sản phẩm theo cấu trúc export. Không copy đề xuất vào mọi dòng biến thể nếu template không yêu cầu.
7. File text/meta không chứa cột ảnh. Nhóm sản phẩm theo tập trường được sửa để tránh cột tùy chọn rỗng xóa dữ liệu của sản phẩm khác.

Ví dụ minh họa header hiện tại, một sản phẩm hai lựa chọn Size, chỉ đổi SEO title và SEO description:

```csv
URL handle,Title,Option1 name,Option1 value,SEO title,SEO description
floral-throw-blanket,Floral Throw Blanket,Size,Small,Floral Throw Blanket with Blue Botanical Pattern,"Explore a floral throw blanket with a blue botanical pattern. View available sizes and product details to choose an option for your space."
floral-throw-blanket,,Size,Large,,
```

Các ô trống dòng tiếp nối ở ví dụ là cấu trúc continuation, không phải chỉ dẫn KEEP/xóa SEO của một sản phẩm độc lập. Phải sao chép cấu trúc thật từ export; không dùng Size Small/Large này cho sản phẩm khác. Nếu export có nhiều option hoặc option liên kết metafield, giữ toàn bộ phụ thuộc đã xác minh; nếu chưa chứng minh được bảo toàn cấu trúc, chuyển phần cập nhật đó sang Matrixify `UPDATE` hoặc thủ công, không đoán.

### 6.3. Tạo file alt ảnh riêng

1. Dùng block dòng ảnh/variant từ export mới của các sản phẩm có alt được duyệt.
2. Giữ Handle, Title và cấu trúc option bắt buộc nguyên trạng; giữ các cột URL ảnh/position cần thiết theo export. Không lấy URL thumbnail từ workbook để thay URL export.
3. Join mỗi `Image_Audit` được duyệt với đúng ảnh gốc trong block sản phẩm. Cập nhật **Image alt text** của dòng đó; các ảnh còn lại giữ giá trị mới nhất trong export.
4. Không thêm URL mới, không đổi Image position, không sửa cột Variant image URL hoặc liên kết ảnh-biến thể; nếu cấu trúc ảnh của export đòi giữ cột liên kết thì giữ nguyên, không sáng tác giá trị.
5. Native import có thể xử lý/re-upload ảnh khi có cột ảnh. Chạy canary và đối chiếu trước/sau ảnh/URL/quan hệ biến thể; nếu có thêm ảnh hoặc đổi quan hệ ngoài kế hoạch, dừng mở rộng và dùng luồng phù hợp hơn. Không hứa alt-only CSV sẽ giữ mọi ID ảnh nếu chưa được kiểm chứng trên shop.

Mapping một dòng ảnh minh họa: `Product image URL=<URL nguyên bản từ export>` và `Image alt text=Blue botanical floral throw blanket draped over a cream sofa`. Không tự tạo URL CDN giả để nhập.

### 6.4. Upload và theo dõi

1. Tạo file canary chứa 3–5 sản phẩm đã được duyệt, gồm loại có nhiều biến thể/ảnh nếu phạm vi có; nếu danh sách ít hơn thì thử số đang có.
2. **Products → Import → Add file**; chọn CSV UTF-8.
3. Bật **Overwrite products with matching handles**.
4. Kiểm tra preview: đúng sản phẩm, đúng trường, đúng số sản phẩm riêng biệt; không suy số sản phẩm từ số dòng CSV. Bất kỳ dấu hiệu tạo mới hoặc khóa lạ nào đều phải dừng trước import.
5. Không bật tùy chọn xuất bản sản phẩm mới; job này không tạo mới. Bấm bước upload/preview và **Import products** theo giao diện hiện tại sau khi kiểm tra xong.
6. Lưu thời điểm, file, manifest và kết quả/thông báo import; đợi hoàn tất trước khi chạy job cùng phạm vi.
7. Re-export và QA 100% trường thay đổi của canary. Chỉ khi đạt mới chạy job tiếp theo, tối đa 20 sản phẩm/job, cùng cơ chế đối chiếu.

Native CSV import đã bắt đầu không thể hủy; preview không phải dry-run kiểm chứng toàn bộ hành vi. Nếu có lỗi, không nhập lại toàn bộ file một cách mù quáng. Đối chiếu sản phẩm thực tế để xác định đã sửa gì trước khi tạo job xử lý tiếp. [Shopify: Import products](https://help.shopify.com/en/manual/products/import-export/import-products).

## 7. Luồng 2B — hàng loạt bằng Matrixify

### 7.1. Export và nhận diện

1. Trong Shopify **Apps → Matrixify**, tạo **New Export**.
2. **Select Sheets → Products**; chọn Basic Columns cần thiết (ID, Handle, Title, Body HTML), Media và Metafields.
3. Với SEO, lọc keys `title_tag`, `description_tag`; tải export sau khi hoàn tất.
4. Dùng ID của chính shop hiện tại. Nếu chỉ có Handle, đối chiếu Handle còn tồn tại và không thay nó. Không dùng ID của shop thử nghiệm cho production.
5. Tạo workbook triển khai riêng có duy nhất sheet **Products**. Chọn một khóa cập nhật chính: mặc định **ID**; không cần gửi Handle khi cập nhật bằng ID để tránh vô tình đổi URL.

Matrixify cung cấp hướng dẫn xuất và cập nhật hai SEO metafield này; cột SEO trống có thể phản ánh fallback từ tên/mô tả sản phẩm. [Matrixify: Bulk update SEO](https://matrixify.app/tutorials/how-to-bulk-update-shopify-seo-title-and-description/).

### 7.2. File Product Title, mô tả và meta

Bắt buộc `Command=UPDATE`. Giá trị này yêu cầu sản phẩm đã tồn tại. `MERGE` có thể tạo mới khi không tìm thấy; không để Command trống vì mặc định là MERGE. Cấm NEW/REPLACE/DELETE trong job SEO. [Matrixify: Products commands](https://matrixify.app/documentation/products/).

Ví dụ sheet **Products**, chỉ đổi hai meta fields; ID dưới đây là giả định, phải thay bằng export thật:

```csv
ID,Command,Metafield: title_tag [string],Metafield: description_tag [string]
1234567890123,UPDATE,Floral Throw Blanket with Blue Botanical Pattern,"Explore a floral throw blanket with a blue botanical pattern. View available sizes and product details to choose an option for your space."
```

Nếu chọn Handle, thay cột ID bằng `Handle` hiện có, giữ Command UPDATE. Nếu duyệt đổi H1 qua Product Title, dùng file chứa khóa + Command + `Title`. Nếu duyệt mô tả, dùng khóa + Command + `Body HTML` với toàn bộ HTML cuối cùng.

Không cần gửi các cột option/variant cho cập nhật Product Title/Body HTML/SEO qua template Matrixify này. Dùng file tối thiểu và bỏ hẳn các cột không cập nhật. [Matrixify: Minimum required columns](https://matrixify.app/tutorials/minimum-columns-to-update-product/).

### 7.3. File alt ảnh

Mỗi ảnh một dòng. Dùng URL gốc do Matrixify export, định danh đúng sản phẩm và ảnh. Ví dụ theo Handle:

```csv
Handle,Command,Image Src,Image Command,Image Alt Text
floral-throw-blanket,UPDATE,<URL_ẢNH_GỐC_1_TỪ_EXPORT>,MERGE,Blue botanical floral throw blanket draped over a cream sofa
floral-throw-blanket,UPDATE,<URL_ẢNH_GỐC_2_TỪ_EXPORT>,MERGE,Close-up of the blue floral pattern on the throw blanket
```

Các placeholder URL phải được thay bằng ảnh hiện có đã kiểm tra. Hai loại Command không được nhầm:

- **Command=UPDATE**: thao tác đối với sản phẩm.
- **Image Command=MERGE**: xử lý media; giữ media khác nhưng có thể thêm media nếu không match.

Matrixify mô tả việc match media theo filename. Vì vậy, dùng nguyên URL export, kiểm tra filename/media đúng trong sản phẩm, không dựa chỉ vào thứ tự. Nếu có filename mơ hồ hoặc shared file cần alt khác nhau, chuyển NEEDS_REVIEW. Không dùng Image Command REPLACE/DELETE, không gửi Image Position trong tác vụ alt-only. [Matrixify: Media behavior](https://matrixify.app/documentation/products/#image-command).

Ảnh bên trong Body HTML không được mặc định xử lý bởi file alt gallery. Muốn sửa phải map và duyệt Description tương ứng. Chạy canary ảnh riêng và yêu cầu số media, thứ tự, liên kết variant và ID nếu quan sát được giữ nguyên; phát hiện thêm media là lỗi phải xử lý trước khi mở rộng.

### 7.4. Dry Run, chạy thật và kết quả

1. Ở Matrixify Home, phần **Import**, upload file.
2. Chờ phân tích, kiểm tra entity **Products**, số sản phẩm riêng biệt, cột được nhận diện và warnings. Không nhập nếu cột SEO không được nhận đúng.
3. Trong Import Options, bật **Dry Run**, chạy để kiểm tra cấu trúc; tải kết quả và giải quyết lỗi.
4. Dry Run không chứng minh Shopify API sẽ chấp nhận mọi dữ liệu, không bảo đảm khóa/ảnh match đúng và không thay thế canary. [Matrixify: Job options](https://matrixify.app/documentation/matrixify-import-export-job-options/).
5. Tạo job thật từ đúng file canary được duyệt, tắt Dry Run, kiểm tra lại **Command=UPDATE**, rồi bấm **Import** khi người vận hành cho phép.
6. Ghi job ID và trạng thái. Nếu đóng trình duyệt hoặc mất kết nối, mở lại job đó; không tạo job mới chỉ vì chưa nhìn thấy kết quả.
7. Khi terminal, tải **Import Results**; đọc trạng thái, **Import Comment**, warnings và đối chiếu actual data. Job Finished không đồng nghĩa mọi dòng thành công. [Matrixify: Import workflow](https://matrixify.app/documentation/import-data-to-shopify-with-matrixify/).
8. Nếu hủy job, các cập nhật đã chạy vẫn có thể đã được ghi. Đối chiếu kết quả và re-export; chỉ retry phần chưa đạt sau khi biết trạng thái thật. Không coi Cancel là rollback.
9. Canary đạt thì triển khai các job còn lại tối đa 20 sản phẩm/job, lưu manifest và QA từng job. Không cho hai job cùng ghi một sản phẩm/ảnh.

## 8. QA kỹ thuật sau triển khai

### 8.1. Đối chiếu dữ liệu lưu trong Shopify — 100% trường thay đổi

1. Re-export tập sản phẩm vừa triển khai bằng cùng nguồn/schema hoặc đọc admin trường tương ứng.
2. Join bằng ID/Handle, với ảnh dùng khóa sản phẩm + URL/media đã đối chiếu.
3. So sánh từng trường SET với approved value; ghi `expected`, `actual`, `PASS/FAIL`, thời điểm và bằng chứng.
4. So sánh các trường bảo vệ và cấu trúc: Handle, Product/Variant ID, option, SKU, giá, media, trạng thái sản phẩm. Sai lệch không giải thích được phải được điều tra, không tự coi là bình thường.
5. Tồn kho có thể đổi do đơn hàng trong lúc shop mở. Kiểm tra payload không ghi inventory và đối chiếu nhật ký nghiệp vụ khi cần; không yêu cầu tồn kho đứng yên hoặc khôi phục tồn kho từ snapshot cũ.
6. Kiểm tra số sản phẩm mới phát sinh do job bằng kết quả và tập khóa trước/sau, không chỉ nhìn tổng catalog nếu shop có hoạt động tạo sản phẩm khác cùng lúc.

`EMPTY` ở SEO field có thể là fallback hoặc do giá trị bằng Title/Description. So sánh cả giá trị lưu và kết quả render trước khi kết luận lỗi. Nếu export không đủ chứng minh, dùng admin/nguồn đọc đã được phép; chưa đủ bằng chứng thì giữ QA chưa hoàn tất.

### 8.2. Kiểm tra HTML/render và SERP Preview

Mở đúng URL, locale và market trong storefront, reload và kiểm tra:

- [ ] Trang trả HTTP 200, không redirect bất ngờ, không 404/soft error.
- [ ] `<title>` phản ánh SEO title và suffix theme đã biết; không lặp tên shop.
- [ ] `<meta name="description">` đúng; không bị app/theme xuất hai giá trị mâu thuẫn.
- [ ] H1 chính đúng Product Title hoặc mapping đã duyệt; không mất tên sản phẩm.
- [ ] Canonical phù hợp, không bị đổi sang sản phẩm khác; không phát sinh noindex/chặn crawl.
- [ ] Description giữ định dạng, bảng size và thông tin quan trọng đã duyệt.
- [ ] Link trong Description đi đúng đích, không gắn UTM vào liên kết nội bộ.
- [ ] Mobile hiển thị ổn, chọn biến thể, nút thêm giỏ và thông tin giá hoạt động. Không cần đặt đơn thanh toán thật để QA nội dung.
- [ ] Nếu theme có Product structured data, tên/giá/tồn kho vẫn khớp nội dung thật; dùng Rich Results Test nếu cần kiểm tra lỗi phát sinh, không thêm schema mới ngoài phạm vi.
- [ ] Xem SERP Preview mobile/desktop để phát hiện title/description khó đọc hoặc bị cắt; không kéo dài/nhồi keyword chỉ để đủ 60/160 ký tự.

Với catalog lớn, kỹ thuật viên có thể tự động thu thập các trường HTML trên 100% URL đã sửa và so sánh manifest; kiểm tra trực quan từng sản phẩm của canary, mỗi template/kiểu biến thể và các trường hợp có diff bất thường. Nếu một loại kiểm tra chưa phủ hết phạm vi, báo đúng coverage thay vì ghi “đã kiểm tra toàn bộ”.

Ví dụ đọc tại DevTools Console của trang đang mở (chỉ đọc DOM, không sửa):

```javascript
({
  title: document.title,
  metaDescriptions: [...document.querySelectorAll('meta[name="description"]')]
    .map(el => el.content),
  h1: [...document.querySelectorAll('h1')].map(el => el.textContent.trim()),
  canonicals: [...document.querySelectorAll('link[rel="canonical"]')]
    .map(el => el.href),
  robots: [...document.querySelectorAll('meta[name="robots"]')]
    .map(el => el.content)
})
```

Đây chỉ là kiểm tra DOM hiện tại; HTTP status và `X-Robots-Tag` cần xem ở Network/response headers. Không có meta robots không tự chứng minh trang indexable.

### 8.3. Alt text và ảnh CDN — kiểm tra đúng lớp dữ liệu

**Alt text không nằm trong file JPG/PNG/WebP của CDN.** Mở trực tiếp CDN URL chỉ kiểm tra ảnh truy cập/hiển thị được; không chứng minh alt đã cập nhật. Alt được lưu trong Shopify và cần hiện ở `<img alt="...">` trên trang. Google cũng dùng nội dung xung quanh ảnh để hiểu ảnh. [Google: Image SEO](https://developers.google.com/search/docs/appearance/google-images).

1. Mở từng media trong Shopify Admin để kiểm tra alt đã lưu, hoặc đối chiếu export sau import.
2. Trên storefront, mở đúng slide/biến thể, cuộn để lazy-load ảnh và Inspect đúng `<img>`.
3. Kiểm tra thuộc tính `alt`, đối chiếu `src`/`srcset`/`currentSrc` để biết đúng ảnh. Responsive URL có tham số kích thước có thể khác URL export; dùng mapping đã xác minh, không đoán theo ảnh thứ N trong DOM.
4. Mở `currentSrc` hoặc xem Network: ảnh phải tải được, không có ảnh mới/trùng hoặc mất liên kết biến thể ngoài kế hoạch.
5. Lặp lại cho ảnh Description nếu được sửa. Nếu admin đúng nhưng HTML sai, kiểm tra theme/app cache và cách render; không upload lại ảnh để cố sửa alt.

Ví dụ lấy dữ liệu ảnh để đối chiếu, kết quả có cả ảnh ngoài gallery nên cần lọc đúng media:

```javascript
[...document.images].map(img => ({
  src: img.getAttribute('src'),
  currentSrc: img.currentSrc,
  srcset: img.getAttribute('srcset'),
  alt: img.getAttribute('alt'),
  loaded: img.complete && img.naturalWidth > 0
}))
```

## 9. Google Search Console và đo lường

### 9.1. Yêu cầu thu thập lại URL

Sau khi QA kỹ thuật đạt:

1. Mở **Google Search Console**, chọn đúng property của domain.
2. Dán URL canonical sản phẩm vào thanh **URL Inspection**.
3. Xem trạng thái index và lần crawl đã ghi nhận; không coi chúng là trạng thái live hiện tại.
4. Chạy **Test live URL** để kiểm tra khả năng truy cập mới. Nếu có lỗi, sửa lỗi thuộc phạm vi hoặc chuyển kỹ thuật viên xử lý trước.
5. Bấm **Request indexing** cho URL ưu tiên đã sửa đáng kể, ghi thời điểm yêu cầu.
6. Không submit lặp lại một URL nhiều lần để cố tăng tốc. Với hàng nghìn URL, dùng sitemap hiện có và liên kết nội bộ; kiểm tra sitemap đã được Search Console biết đến, không gửi từng URL vô hạn.

Đây là **yêu cầu** Google thu thập lại, không phải lệnh bắt buộc. Google nêu có thể mất vài ngày đến vài tuần; không bảo đảm index hay xếp hạng ngay. Request indexing có quota và gửi lại nhiều lần không giúp nhanh hơn. [Google: Request recrawl](https://developers.google.com/search/docs/crawling-indexing/ask-google-to-recrawl).

### 9.2. Đo lường theo URL và nhóm truy vấn

- Trước triển khai: xuất dữ liệu Search Console theo Page, Query, Country, Device với kỳ tham chiếu được ghi rõ, ví dụ 28 ngày trước ngày sửa.
- Ghi `deployed_at`, phạm vi trường đổi, revision và nhóm keyword mục tiêu. Phân biệt keyword cũ với keyword mới.
- Kiểm tra sau 7/14/28 ngày như mốc vận hành: crawl, impressions, clicks, CTR, vị trí trung bình theo cùng bộ lọc. Dữ liệu quá ít hoặc có độ trễ phải ghi “chưa đủ dữ liệu”.
- Với shop có GA4/purchase tracking đã kiểm tra: theo dõi Google organic landing pages, đơn và doanh thu; không khẳng định nối được từng keyword organic với từng đơn.
- So sánh với sản phẩm tương đồng không sửa nếu có; ghi nhận khuyến mãi, thay đổi giá và mùa vụ. Không suy ra mọi tăng/giảm đều do title.
- Google vẫn hiện title/snippet cũ ngay sau deploy không phải bằng chứng import thất bại. Nghiệm thu dữ liệu Shopify/HTML trước, theo dõi Google xử lý sau.

Search Console và Analytics có mục đích và phương pháp đo khác nhau; cần đối chiếu theo phạm vi tương ứng. [Google: Combining Search Console and Analytics](https://developers.google.com/search/docs/monitor-debug/google-analytics-search-console).

## 10. Rollback và xử lý lỗi

### 10.1. Khi nào dừng mở rộng

- Sai sản phẩm hoặc sai ảnh được cập nhật.
- Phát sinh sản phẩm/media mới, mất variant hoặc thay Handle ngoài kế hoạch.
- Trường ngoài danh sách cho phép bị đổi do job.
- Payload không đúng revision duyệt, giá trị rỗng xóa dữ liệu, HTML/H1 vỡ.
- Có lỗi chưa rõ phạm vi; không tiếp tục hàng loạt để “xem có tự hết không”.

### 10.2. Các bước khôi phục

1. Dừng job tiếp theo. Với Matrixify đang chạy, có thể Cancel rồi chờ trạng thái cuối và tải kết quả; với native CSV, không có Cancel sau khi bắt đầu.
2. Lưu trạng thái hiện tại và kết quả import trước khi sửa tiếp.
3. Xác định chính xác khóa và trường đã bị tác động; export lại để biết phần thành công/thất bại.
4. So sánh giá trị hiện tại với giá trị job vừa ghi. Nếu người khác đã sửa tiếp, cần xử lý xung đột; không đè bằng backup cũ.
5. Tạo payload rollback **chỉ các trường đã thay đổi** từ snapshot gốc, cùng khóa hiện có. Không re-import toàn bộ backup sản phẩm chứa giá/kho cũ.
6. Với Matrixify vẫn dùng Command UPDATE; với native vẫn giữ cấu trúc option và khóa. Khôi phục alt dùng đúng media còn tồn tại, không tải ảnh mới để che lỗi mapping.
7. Nếu SEO field gốc là EMPTY/fallback: đây là thao tác xóa override có chủ đích, tách riêng khỏi job SET thông thường. Thử trên một sản phẩm; ưu tiên thao tác admin được kiểm tra hoặc cách xóa metafield theo tài liệu của công cụ hiện tại. Không tự điền nội dung render cũ và gọi đó là khôi phục nguyên trạng. Nếu không xác minh được fallback đã trở lại, đánh dấu rollback chưa hoàn tất.
8. QA lại admin/export, HTML và các trường bảo vệ; ghi `ROLLED_BACK` trong nhật ký triển khai, không thêm giá trị này vào review_status.

Nếu variant/media đã bị xóa và tạo lại, CSV không bảo đảm phục hồi đúng ID và các liên kết app cũ. Chuyển kỹ thuật viên xử lý, ghi phạm vi ảnh hưởng; không tuyên bố rollback hoàn toàn chỉ vì tên/ảnh nhìn giống trước.

## 11. Ngoại lệ đổi URL Handle và 301

Mặc định **không đổi Handle**. Chỉ thực hiện khi có chỉ định riêng bằng văn bản cùng danh sách URL cũ/mới, người duyệt và kế hoạch kiểm tra redirect. Không đưa ngoại lệ này vào job SEO text thông thường.

Ví dụ kế hoạch:

```text
Old: /products/old-floral-blanket
New: /products/floral-throw-blanket
Expected: Old trả 301, Location trỏ New; New trả 200.
```

Nếu đã được phép: sửa đúng URL handle trong Search engine listing; giữ tùy chọn tạo redirect nếu giao diện cung cấp, hoặc vào **Content → Menus → URL redirects → Create URL redirect**, nhập đường dẫn cũ và đích mới rồi lưu. Đối chiếu redirect đã tồn tại trước khi tạo trùng. Shopify redirects chỉ hoạt động trong những điều kiện nhất định, ví dụ đường dẫn nguồn không còn trang hoạt động. [Shopify: URL redirects](https://help.shopify.com/en/manual/online-store/menus-and-links/url-redirect).

Kỹ thuật viên kiểm tra HTTP 301/Location thật, không chỉ thấy trình duyệt mở được đích; tránh chuỗi redirect, vòng lặp, đích 404. Cập nhật liên kết nội bộ/canonical/sitemap liên quan theo kế hoạch. Nếu rollback Handle, kiểm tra cả redirect cũ và mới để không tạo vòng lặp. Không dùng native CSV với Handle mới như cách “đổi slug”, vì nó có thể được coi là sản phẩm mới.

## 12. Checklist bàn giao và vận hành ở quy mô lớn

Mỗi job có manifest gồm `shop_domain`, `job_id`, `batch_id`, `source_export_ref`, `source_exported_at`, danh sách khóa, trường/giá trị before-after, người duyệt, revision, thời điểm triển khai, job/result file và QA result. Giữ bản bất biến của file đã thực sự import.

- [ ] Backup Shopify CSV đã tải và mở kiểm tra; nguồn ID/media bổ sung đã lưu nếu cần.
- [ ] Không còn khóa không match, trùng khóa sản phẩm trong sheet chính hoặc mâu thuẫn ID/Handle.
- [ ] Chỉ APPROVED đúng revision và approved_fields có mặt trong manifest.
- [ ] Có báo cáo QA nội dung đúng bản/revision cho sản phẩm được triển khai; lỗi ảnh hưởng trường triển khai đã xử lý, mọi giới hạn được người duyệt ghi nhận rõ. QA_PASS không thay cho APPROVED.
- [ ] Payload qua allowlist; không có giá, SKU, kho, shipping, published, tags hoặc dữ liệu ngoài phạm vi.
- [ ] Ảnh được map đúng; không dùng thứ tự gallery làm định danh duy nhất.
- [ ] File native giữ option/variant; file Matrixify đúng sheet Products và Command UPDATE.
- [ ] Canary đã qua đối chiếu dữ liệu và kiểm tra trực quan; không dùng Dry Run thay canary.
- [ ] Mỗi job có trạng thái terminal và kết quả đã đọc trước khi retry.
- [ ] Đối chiếu 100% trường thay đổi, ghi coverage HTML/ảnh và mọi phần chưa kiểm tra.
- [ ] Không có sản phẩm/media mới hoặc mất biến thể do job; sai lệch ngoài phạm vi đã được giải thích/xử lý.
- [ ] URL/Handle, canonical và khả năng truy cập giữ đúng; không có lỗi thêm giỏ liên quan thay đổi.
- [ ] Đã lưu kế hoạch/nguồn rollback và nhật ký Search Console nếu có submit.
- [ ] Chỉ đánh dấu hoàn tất khi mọi job trong phạm vi được phép đều có bằng chứng QA; phần chặn báo PARTIAL.

Với hàng nghìn sản phẩm, tự động hóa việc join, lọc trạng thái, sinh file, kiểm tra allowlist, đếm khóa và so sánh before-after. Không tự động hóa bằng cách giảm việc xem ảnh hoặc tự duyệt các đề xuất chưa được kiểm tra. Dữ liệu hồ sơ là nguồn chính; workbook triển khai được sinh lại từ đó để tránh chép tay sai cột.

## 13. Cách dùng ba tài liệu trong dự án

1. Điền URL, thị trường, ngôn ngữ và phạm vi trong `prompt.md`; cung cấp export quản trị khi có để đối chiếu trường gốc.
2. Cho Agent chạy kiểm kê và lô đầu 10 sản phẩm bằng `prompt.md`, lưu workbook trong `resutls`; xác nhận riêng từng lô nghiên cứu tiếp theo.
3. Giao `prompt_qa.md` cùng workbook và danh sách sản phẩm để kiểm tra độc lập từng lô 10 sản phẩm. Đọc điểm sản phẩm/ảnh, độ phủ và lỗi trong báo cáo QA; xác nhận riêng từng lô QA tiếp theo. Có thể dùng phiên làm việc khác cho QA, nhưng việc mở lại nguồn và ảnh vẫn bắt buộc theo prompt QA.
4. Sửa đề xuất theo lỗi đã kiểm chứng, cập nhật revision và QA lại. Duyệt nội dung theo revision; xác nhận tiếp lô hoặc điểm QA cao không thay cho bước duyệt này.
5. Người vận hành dùng SOP để tạo backup và payload từ các trường APPROVED. Agent tạo file không được tự suy ra quyền ghi lên Shopify.
6. Chạy luồng thủ công hoặc bulk, QA sau triển khai, theo dõi và lưu nhật ký. Trước một đợt mới, đối chiếu lại header/export và tài liệu nguồn vì Shopify, theme và Matrixify có thể thay đổi.

Tài liệu này là quy trình đã đối chiếu nguồn, không phải báo cáo đã import hoặc thử thành công trên một shop cụ thể. Đến khi có dữ liệu và quyền triển khai thật, các bước canary/QA vẫn là điều kiện bắt buộc của người vận hành.
