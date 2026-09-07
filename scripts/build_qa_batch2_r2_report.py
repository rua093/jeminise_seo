from pathlib import Path
import build_qa_batch2_report as base
import json, hashlib

ROOT = Path(__file__).resolve().parents[1]
base.QA_RUN_ID = "20260907_123726"
base.QA_BATCH_ID = "qa_batch_002_r2"
base.R2_MODE = True
base.QA_DIR = ROOT / "seo_runs" / "jeminise.com" / "20260906_234129" / "qa" / base.QA_RUN_ID
base.OUT_DIR = ROOT / "resutls" / "jeminise.com" / "20260906_234129" / "qa" / base.QA_RUN_ID
base.SOURCE = ROOT / "resutls" / "jeminise.com" / "20260906_234129" / "revisions" / "qa_batch_002_r2" / "SEO_Product_Optimization_qa_batch_002_r2.xlsx"
base.REASONS["K1"] = "Đối chiếu long-tail với thiết kế riêng, nhóm Christian/Christmas và mức độ trùng intent giữa các trang cùng cụm."

def ensure_checkpoint_files():
    base.QA_DIR.mkdir(parents=True, exist_ok=True)
    src_hash = hashlib.sha256(base.SOURCE.read_bytes()).hexdigest()
    manifest = base.QA_DIR / "qa_manifest.json"
    if not manifest.exists():
        manifest.write_text(json.dumps({
            "qa_run_id": base.QA_RUN_ID, "qa_batch_id": base.QA_BATCH_ID, "revision": "r2",
            "shop": "jeminise.com", "source_run_id": "20260906_234129",
            "source_workbook": str(base.SOURCE), "source_sha256_at_freeze": src_hash,
            "frozen_at": base.common.now(), "prompt_qa": str(ROOT/"seo-prompt/jeminise/prompt_qa.md"),
            "prompt": str(ROOT/"seo-prompt/jeminise/prompt.md"), "status": "IN_PROGRESS"
        }, ensure_ascii=False, indent=2), encoding="utf-8")
    progress = base.QA_DIR / "qa_progress.json"
    if not progress.exists():
        progress.write_text(json.dumps({"qa_run_id": base.QA_RUN_ID, "qa_batch_id": base.QA_BATCH_ID,
            "revision": "r2", "current_stage": "SOURCE_RECHECK", "completed_image_keys": [],
            "awaiting_confirmation": False}, ensure_ascii=False, indent=2), encoding="utf-8")

if __name__ == "__main__":
    ensure_checkpoint_files()
    base.main()
    report = base.OUT_DIR / "SEO_QA_qa_batch_002_r2.md"
    txt = report.read_text(encoding="utf-8")
    txt = txt.replace("# SEO QA — qa_batch_002", "# SEO Re-QA — qa_batch_002_r2")
    txt = txt.replace("- Phạm vi: **10 sản phẩm, 64/64 ảnh (100%)**; chỉ inventory position 11–20.", "- Phạm vi: **10 sản phẩm, 64/64 ảnh (100%)**; chỉ inventory position 11–20; revision **r2**.")
    old = "1. **CRITICAL — products 11–16:** draft dùng claim `Personalized`, nhưng purchase options live không có input tên/verse/birth flower; phải chứng minh control hoạt động hoặc bỏ claim.\n2. **CRITICAL — product 19:** description ghi `Customization: 1 text input`, nhưng trang live chỉ có size và pillowcase quantity.\n3. **MAJOR — cả 10 sản phẩm:** description HTML chứa câu nội bộ yêu cầu review, chưa publish-ready.\n4. **MAJOR — ảnh:** alt/observation theo mẫu vị trí sai hàng loạt; nhóm Christian sai nhiều ảnh 3–8, nhóm Christmas thường đảo ảnh cận cảnh với ảnh sham.\n5. **MAJOR — products 11–16:** SERP evidence đã nộp chủ yếu là blanket trong khi keyword nhắm comforter/bedding, đồng thời sáu trang cạnh tranh intent rất gần nhau.\n6. **MAJOR — product 13:** draft chép `Style: Christian Knight Templar`, mâu thuẫn với thiết kế God Is Within Her dành cho nữ đang hiển thị."
    new = "1. **MAJOR — toàn bộ 64 ảnh:** nhiều observation/alt r2 vẫn gán theo vị trí chung hoặc đảo cận cảnh–sham; cần sửa theo nội dung ảnh thực tế.\n2. **MAJOR — 10 sản phẩm:** r2 đã bỏ câu nội bộ draft/QA nhưng còn quá generic và chưa nêu đầy đủ dữ kiện đã xác minh.\n3. **MAJOR — products 11–20:** live Customizer có trường tên (bắt buộc ở 11–16, tùy chọn ở 17–20), nhưng copy r2 chưa ánh xạ rõ điều kiện personalization.\n4. **MAJOR — products 11–16:** SERP còn lệch intent (blanket so với comforter/bedding) và các trang cùng cụm cạnh tranh quá gần nhau.\n5. **MAJOR — product 13:** cụm `Christian Knight Templar` vẫn mâu thuẫn với thiết kế God Is Within Her đang hiển thị.\n6. **MINOR — một số title/H1 nguồn:** còn ký tự `�` hoặc từ `Bedsprea` bị cắt; cần xác nhận bản live/admin trước khi sửa."
    txt = txt.replace(old, new)
    report.write_text(txt, encoding="utf-8")
