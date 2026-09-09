# SEO Product Optimization Revision - qa_batch_032_r5

Scope: batch 032, products 311-320.

Source revision: `qa_batch_032_r4`.

QA source: `resutls/jeminise.com/20260906_234129/qa/20260909_132406/SEO_QA_qa_batch_032_r4.md`.

Changes made:
- Revised only the 4 non-passing products `311-314`.
- Confirmed the best public SEO product type should be `Quilt`/`Quilt Set`, not `Comforter`, because live `.js` product type, Shopify admin export `Type`, size option labels, current admin SEO title, and contact-sheet info panels all support quilt terminology.
- Updated `title_proposed`, `meta_title_seo`, `meta_description_seo`, `description_proposed_html`, `primary_keyword`, `secondary_keywords`, and `meta_keyword` for those 4 products.
- Updated all 32 related `Image_Audit` rows so observations and alt text use `quilt` instead of `comforter`.
- Updated related `Keyword_Map`, `Product_Evidence`, and `Buyer_Search_Research` wording for consistency.
- Preserved the 6 products that already passed QA.
- No `APPROVED` status was created.

QA focus for next check:
- Recheck that products `311-314` no longer trigger product-type consistency MAJOR.
- Recheck all 70 images in batch 032 and confirm no alt still uses `comforter` for the deer quilt products.
