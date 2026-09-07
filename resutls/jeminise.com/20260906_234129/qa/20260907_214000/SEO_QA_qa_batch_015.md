# SEO QA — qa_batch_015

## Kết luận

- Phạm vi: **10 sản phẩm, 57/57 ảnh (100%)**; chỉ inventory position 141–150.
- Điểm lô: **62.8/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: 10 QA_FAIL, 0 QA_REVISE, 0 QA_PASS.
- Phát hiện: 0 CRITICAL, 35 MAJOR, 46 MINOR, 10 LIMITATION.
- Workbook nguồn: `D:\Shopify_Workspace\jeminise_seo\resutls\jeminise.com\20260906_234129\batches\SEO_Product_Optimization_through_batch_015.xlsx`
- SHA-256 lúc đóng băng và bàn giao: `0311DF4CA9BB483107A3267536E7AD4632CCB24F7A0178BF2D9DAD1BA7E6333D`
- Note: this is the complete independent QA following `prompt_qa.md`, revision r1.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
| 141 | Fantasy Fire Dragon Quilt Set with Glowing Lava And Bedroom Setting | 62.5 | QA_FAIL | 0/4/4/1 |
| 142 | Fantasy Mythical Dragon Quilt Set with Winged Dragon And Bedroom Setting | 62.5 | QA_FAIL | 0/4/4/1 |
| 143 | Fantasy Nature Tree of Life Quilt Set with Entwined Golden Roots And Bedroom Setting | 62.5 | QA_FAIL | 0/3/6/1 |
| 144 | Fantasy Nature Tree of Life Quilt Set with Twisted Tree Trunk And Bedroom | 62.5 | QA_FAIL | 0/3/5/1 |
| 145 | Fantasy Quilt Set with Dragon and Bedroom and Fire And Spikes | 62.5 | QA_FAIL | 0/4/4/1 |
| 146 | Fantasy Winged Dragon Quilt Set with Starry Nebula Pattern And Bedroom Setting | 62.5 | QA_FAIL | 0/4/4/1 |
| 147 | Fantasy Winged Dragon Quilt Set with Volcanic Eruption And Bedroom | 65.0 | QA_FAIL | 0/4/3/1 |
| 148 | Floral Nature Cardinal Within Ornate Oval Frame Patchwork Winter Quilt | 62.5 | QA_FAIL | 0/3/4/1 |
| 149 | Folklore Tree of Life in Celtic Knot Quilt | 62.5 | QA_FAIL | 0/3/6/1 |
| 150 | Fox Patchwork Pattern Quilt Set, Animal Quilt Set Patchwork Style | 62.5 | QA_FAIL | 0/3/6/1 |

## Lỗi ưu tiên

1. **MAJOR — product 141:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
2. **MAJOR — product 141:** A publishable SEO title/H1 pair is incomplete when the H1 field is blank. Đề xuất: Write one English H1 that matches the verified product type and visible design; do not copy unsupported claims.
3. **MAJOR — product 141:** Generic audience wording can be irrelevant or contradictory to the visible motif and is not evidence-led storefront copy. Đề xuất: Replace it with concise English customer-facing copy tied only to the verified motif, product type, size/pillowcase choices and supported customizer control.
4. **MAJOR — product 141:** Batch cluster contains several very similar products, so keyword/intent differentiation must be sharper. Đề xuất: Add unique visible motif modifiers and avoid assigning the same generic product keyword to multiple URLs.
5. **MAJOR — product 142:** Internal workflow language is not publish-ready storefront copy. Đề xuất: Remove internal QA/import block and replace with customer-facing verified copy.
6. **MAJOR — product 142:** A publishable SEO title/H1 pair is incomplete when the H1 field is blank. Đề xuất: Write one English H1 that matches the verified product type and visible design; do not copy unsupported claims.
7. **MAJOR — product 142:** Generic audience wording can be irrelevant or contradictory to the visible motif and is not evidence-led storefront copy. Đề xuất: Replace it with concise English customer-facing copy tied only to the verified motif, product type, size/pillowcase choices and supported customizer control.
8. **MAJOR — product 142:** Batch cluster contains several very similar products, so keyword/intent differentiation must be sharper. Đề xuất: Add unique visible motif modifiers and avoid assigning the same generic product keyword to multiple URLs.

## SERP và keyword

Mỗi sản phẩm đã được ghi lại tối thiểu primary keyword và một comparator gần nhất trong `serp_evidence.json`. Evidence chỉ là US public SERP/product comparables, không có claim volume.

## Giới hạn và trạng thái bàn giao

- Không có Shopify admin export; không suy ra admin SEO fields hoặc alt hiện tại ngoài storefront/product.js.
- Đã dùng workbook snapshot, live_source_comparison, product.js/static HTML, customizer_audit và ảnh tải trực tiếp trong thư mục QA run.
- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.
- **XLSX: COMPLETE.** Workbook QA has exactly five sheets, auditable formulas, filters, frozen headers, wrapped text and hyperlinks; structure, formulas and layout were reopened and checked.
- Chưa QA batch tiếp theo sau qa_batch_015. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA data: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_214000\qa_dataset.json`
- SERP evidence: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_214000\serp_evidence.json`
- Validation: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_214000\validation_results.json`
- Manifest/checkpoint: `D:\Shopify_Workspace\jeminise_seo\seo_runs\jeminise.com\20260906_234129\qa\20260907_214000`
