# Git Change Audit - Tool_insert_excel

Ngay audit: 2026-07-16

Pham vi da kiem tra:

- Commit local chua push: `git log origin/main..HEAD`, `git diff origin/main..HEAD`
- File da stage: `git diff --cached`
- File da sua nhung chua stage: `git diff`
- File moi chua duoc Git theo doi: `git ls-files --others --exclude-standard`

Luu y quan trong: khong gop commit local chua push voi Working Tree. Cac nhan dinh duoi day dua tren Git metadata, diff co chon loc, ten route/function/class va doc sau cac cum thay doi co rui ro.

## 1. Tong quan

- Tong so file thay doi theo 4 lop audit: 295
- Commit local chua push: 1 commit, 4 file
- File Modified: 14
  - 2 file Modified trong commit local chua push: `.dockerignore`, `.gitignore`
  - 12 file Modified trong Working Tree: `.env.example`, `Dockerfile`, `README.md`, `docx_template_builder_ver0.9.2.py`, `generation_service.py`, `job_worker.py`, `scripts/deploy_web.sh`, `storage_backend.py`, `templates/configure.html`, `templates/index.html`, `templates/job_status.html`, `web_app.py`
- File Untracked: 279
- File Added: 2 trong commit local chua push: `ulnl_tool/resources/ULNL_template_base.xlsx`, `ulnl_tool/resources/viettel_logo.png`
- File Deleted: 0
- File da stage: 0
- So file source code: 18
- So file cau hinh / deploy: 6
- So file giao dien: 8
- So file tai lieu: 2
- So file tam / generated / upload / binary bundle: 260 untracked trong `tools/LibreOfficePortableLibreOfficePortable/`; ngoai ra co nhieu file ignored nhu `.venv/`, `sessions/`, `__pycache__/`, upload Word/Excel, log local
- So nhom thay doi thuc te: 10

Thong ke nhanh theo lop:

| Lop Git | So file | Them / xoa dong | Ghi chu |
|---|---:|---:|---|
| Local commits `origin/main..HEAD` | 4 | 12 dong text + 2 binary | Commit `ab2a865 Include ULNL template resources in deploy` |
| Staged | 0 | 0 | `git diff --cached` rong |
| Working tree modified | 12 | 6356 them / 297 xoa | Nhieu thay doi lon o web, DOCX engine, UI |
| Untracked | 279 | Chua tracked | 19 file code/UI/docs + 260 file LibreOffice portable |

Phan loai nhanh theo nhom A-N:

- A. Chuc nang nghiep vu moi: CSDL tool, ULNL tool, Data Zone builder, multi-Excel, preview PDF, TOC settings.
- B. Sua loi: header row scoring, duplicate upload filename, NA placeholder detection, local job flow thay vi generate dong bo.
- C. Giao dien HTML/CSS/JS: `templates/index.html`, `configure.html`, `job_status.html`, `csdl_tool.html`, `csdl_job_status.html`, `data_zone_tool.html`, `ulnl_tool.html`, `tktkl.html`.
- D. Refactor/tai cau truc: tach `generation_service`, them `env_loader`, them local background job, session manifest nhieu Excel.
- E. Cau hinh moi truong: `.env.example`, `.gitignore`, `.dockerignore`, `env_loader.py`.
- F. Docker/Cloud/GCP: `Dockerfile`, `scripts/deploy_web.sh`, `storage_backend.py`, `job_worker.py`.
- G. Dependency: them yeu cau LibreOffice system/package; `requirements.txt` khong doi.
- H. Database/data model: `tk_csdl_tool.py`, `data_zone_builder.py`, `ulnl_tool/core/data_model.py`.
- I. Test: khong thay file test moi.
- J. Tai lieu: `README.md`, `UI_GUIDELINES.md`.
- K. File sinh tu dong/tam/output: `tools/LibreOfficePortableLibreOfficePortable/`, ignored `.venv/`, `sessions/`, `__pycache__/`, log, upload Word/Excel.
- L. File upload cua nguoi dung: cac file Word/Excel upload dang ignored, khong nam trong untracked do `.gitignore`.
- M. Nguy co bi mat: chua thay secret/key/token ro rang trong untracked/source; co keyword false positive nhu `token` trong code parsing va `help.key_` cua LibreOffice help.
- N. Chua xac dinh: encoding hien thi mojibake trong output terminal cua mot so file moi can xac minh bang browser/editor.

## 2. Bang quyet dinh

| ID | Nhom thay doi | File chinh | Y nghia | Anh huong | Rui ro | De xuat | Do tin cay |
|----|---------------|-----------|---------|-----------|--------|---------|-------------|
| L1 | Commit local chua push: mo ignore de deploy tai nguyen ULNL | `.gitignore`, `.dockerignore`, `ulnl_tool/resources/ULNL_template_base.xlsx`, `viettel_logo.png` | Cho phep commit/deploy template Excel va logo cho ULNL trong khi van ignore Excel/upload khac | Anh huong Git va Docker build context | Trung binh: `.gitignore` chua ignore dung thu muc portable dang untracked | KEEP_AFTER_FIX | Cao |
| G1 | Bien trang chu thanh Solution Hub nhieu module | `templates/index.html`, `templates/tktkl.html`, route `/tktkl`, `/csdl`, `/ulnl`, `/data-zone` trong `web_app.py` | Bien app tu mot tool DOCX thanh hub gom nhieu cong cu | Doi entrypoint UI, nguoi dung khong vao thang upload DOCX nua | Trung binh: thay doi san pham/UX can xac nhan | NEED_DECISION | Cao |
| G2 | Preview PDF bat buoc truoc khi download DOCX | `preview_service.py`, `web_app.py`, `job_worker.py`, `templates/job_status.html`, `Dockerfile`, `.env.example`, `README.md` | Sau khi generate DOCX, tao PDF preview bang LibreOffice; user xem preview truoc khi tai DOCX | Local can LibreOffice; Docker cai LibreOffice; cloud upload preview len GCS | Cao: download bi chan neu preview chua xem; LibreOffice co the fail/khac Word | KEEP_AFTER_FIX | Cao |
| G3 | Multi-Excel va Step 6 nang cao cho DOCX builder | `generation_service.py`, `web_app.py`, `templates/configure.html`, `docx_template_builder_ver0.9.2.py` | Mot session co the co nhieu Excel; chon nguon theo heading, split/merge, hyperlink noi bo, import/export setup | Tang nang luc generate tai lieu phuc tap | Cao: diff rat lon, can regression test nhieu mau DOCX/Excel | KEEP_AFTER_FIX | Trung binh |
| G4 | Cau hinh muc luc TOC va cap nhat field khi mo Word | `docx_template_builder_ver0.9.2.py`, `generation_service.py`, `templates/configure.html`, `preview_service.py` | Them buoc TOC styles, bookmark vung noi dung, danh dau TOC dirty/update-on-open | Anh huong truc tiep format DOCX output | Cao: OOXML phuc tap, LibreOffice/Word render co the lech | KEEP_AFTER_FIX | Trung binh |
| G5 | Them tool CSDL/DQC/STD | `tk_csdl_tool.py`, `templates/csdl_tool.html`, `templates/csdl_job_status.html`, route CSDL trong `web_app.py` | Upload DOCX thiet ke CSDL, parse bang/field, gan rule DQC/STD, xuat Excel | Them workflow nghiep vu moi doc lap | Trung binh: code lon, chua co test, co bootstrap dependency check khi import | NEED_DECISION | Trung binh |
| G6 | Them tool Data Zone Raw/Work | `data_zone_builder.py`, `templates/data_zone_tool.html`, route `/data-zone/*` | Tao phu luc Raw Zone/Work Zone tu Excel nguon | Them workflow Excel moi doc lap | Trung binh: can file mau de xac nhan mapping cot | NEED_DECISION | Trung binh |
| G7 | Them tool ULNL | `ulnl_tool/core/*`, `ulnl_tool/scripts/*`, `templates/ulnl_tool.html`, resources trong local commit, route `/ulnl/*` | Import danh sach chuc nang, classify ma CN, build file ULNL Excel va preview PDF | Them workflow nghiep vu lon, phu thuoc template/resource va LibreOffice de preview | Cao: resources da commit nhung code untracked; can quyet dinh commit ca bo hay khong | NEED_DECISION | Trung binh |
| G8 | Cai thien deploy Cloud Run/GCS | `Dockerfile`, `scripts/deploy_web.sh`, `storage_backend.py`, `job_worker.py` | Cai LibreOffice trong image, them preview object GCS, doi default service web thanh `solutionhub` | Anh huong Docker image size, Cloud Run service name, worker manifest | Trung binh: doi ten service co the tao service moi hoac deploy sai target | KEEP_AFTER_FIX | Cao |
| G9 | Tai lieu/UI guideline | `README.md`, `UI_GUIDELINES.md` | Ghi huong dan preview LibreOffice va chuan UI wizard | Ho tro dev/operating | Thap: `README` can them thong tin cho cac module moi neu giu | KEEP | Cao |
| G10 | LibreOffice portable bung vao repo | `tools/LibreOfficePortableLibreOfficePortable/**` 260 file, 4.1 MB | Ban copy local cua LibreOffice Portable de convert preview | Khong nen la source; khong khop ignore pattern hien tai | Cao: binary/tool bundle, temp/config, co the tang repo va sai license/distribution | IGNORE_FROM_GIT | Cao |

## 3. Chi tiet tung nhom

### L1. Commit local chua push: mo ignore de deploy tai nguyen ULNL

1. Muc dich: cho phep Git va Docker mang theo `ulnl_tool/resources/ULNL_template_base.xlsx` va `viettel_logo.png`, trong khi van ignore Excel/Word/upload noi chung.
2. File chinh:
   - `.gitignore`: Modified trong commit local.
   - `.dockerignore`: Modified trong commit local.
   - `ulnl_tool/resources/ULNL_template_base.xlsx`: Added trong commit local.
   - `ulnl_tool/resources/viettel_logo.png`: Added trong commit local.
3. Truoc thay doi: cac file `.xlsx/.png` resources co the bi ignore/khong vao Docker context tuy pattern.
4. Sau thay doi: resources ULNL duoc phep commit/deploy.
5. Co duoc su dung: co. `web_app.py` khai bao `ULNL_TEMPLATE_PATH = BASE_DIR / "ulnl_tool" / "resources" / "ULNL_template_base.xlsx"` va import logic build ULNL.
6. Dependency/cau hinh moi: khong them Python dependency, nhung phu thuoc file binary resource ton tai.
7. Anh huong local: neu commit L1 da co, local co template/logo cho ULNL.
8. Anh huong Docker/Cloud: Docker build co the dua resource vao image do `.dockerignore` co exception.
9. Rui ro loi: thap-trung binh. Exception ignore hop ly, nhung neu resource khong duoc cap nhat cung code ULNL thi route ULNL loi.
10. Rui ro bao mat: thap, day la Excel template va PNG; khong thay secret.
11. Migration/cai dat/env: khong.
12. Kiem thu: build Docker image va vao `/ulnl`, generate file mau toi buoc build.
13. Code thua/chua hoan thien: `.gitignore` commit nay ignore `tools/libreoffice/` va `tools/LibreOfficePortable/`, nhung untracked hien tai la `tools/LibreOfficePortableLibreOfficePortable/`, chua bi ignore.
14. De xuat: KEEP_AFTER_FIX. Giu commit, nhung bo sung ignore pattern dung cho thu muc portable phat sinh truoc khi push commit tiep theo.
15. Do tin cay: Cao.

### G1. Bien trang chu thanh Solution Hub nhieu module

1. Muc dich: thay vi app mot luong DOCX, trang chu tro thanh hub 4 module: Template, CSDL, ULNL, Data Zone.
2. File chinh:
   - `templates/index.html`: Modified.
   - `templates/tktkl.html`: Untracked.
   - `web_app.py`: Modified, them route `tktkl_index`, `csdl_index`, `data_zone_index`, `ulnl_index`.
3. Truoc thay doi: `/` hien form upload Word template va Excel de vao configure DOCX.
4. Sau thay doi: `/` chi la man hinh chon module; upload DOCX chuyen sang `/tktkl`.
5. Co duoc su dung: co. Route trong `web_app.py` tro ve cac template moi.
6. Dependency/cau hinh moi: khong.
7. Anh huong local: nguoi dung quen workflow cu phai bam them 1 buoc.
8. Anh huong Docker/Cloud: anh huong UI production; footer co hien URL `solutionhub-bitaaxnokq-as.a.run.app`, co the la hard-code theo moi truong.
9. Rui ro loi: trung binh. Neu `templates/tktkl.html` chua dong bo voi `/configure`, module Template co the hong luong upload.
10. Rui ro bao mat: thap, nhung hard-code URL production co the lam lo endpoint noi bo hoac gay nham moi truong.
11. Migration/cai dat/env: khong.
12. Kiem thu: mo `/`, `/tktkl`, upload 1 DOCX + 1 Excel, dam bao vao `/configure`.
13. Code thua/chua hoan thien: can xac minh `tktkl.html`; hien tai la untracked nen neu deploy ma quen commit se loi link.
14. Quyet dinh: NEED_DECISION. Day la thay doi san pham/UX, can xac nhan co muon doi app thanh hub hay giu tool DOCX lam trang chinh.
15. Do tin cay: Cao.

### G2. Preview PDF bat buoc truoc khi download DOCX

1. Muc dich: sau generate, app tao PDF tu DOCX bang LibreOffice va yeu cau user xem preview truoc khi tai file DOCX.
2. File chinh:
   - `preview_service.py`: Untracked, tim LibreOffice, update TOC page numbers bang Word COM tren Windows, convert DOCX sang PDF.
   - `web_app.py`: Modified, route `/jobs/{job_id}/preview`, local background job, chan download neu preview chua xem.
   - `job_worker.py`: Modified, cloud worker tao/upload preview PDF.
   - `storage_backend.py`: Modified, them `job_preview_object`.
   - `templates/job_status.html`: Modified, iframe preview va enable download sau khi load.
   - `Dockerfile`: Modified, cai `libreoffice`, `libreoffice-writer`, fonts.
   - `.env.example`, `README.md`: Modified, huong dan `LIBREOFFICE_BIN`.
3. Truoc thay doi: local generate tra thang FileResponse DOCX; cloud job chi output DOCX.
4. Sau thay doi: local/cloud deu tao job, convert preview PDF, user xem preview roi moi download DOCX.
5. Co duoc su dung: co. `web_app.py` import `convert_docx_to_pdf`; `job_worker.py` import khi chay worker.
6. Dependency/cau hinh moi:
   - Local Windows: can LibreOffice hoac `LIBREOFFICE_BIN`.
   - Docker/Cloud: image cai LibreOffice qua apt.
   - Tien trinh Windows co the dung Microsoft Word COM neu co Word.
7. Anh huong local: can LibreOffice; neu khong co, preview fail nhung code cho phep enable download khi `preview_error` o UI, trong khi backend `/jobs/{job_id}/download` chi chan khi `preview_available` true va chua seen.
8. Anh huong Docker/Cloud: image nang hon, build lau hon; worker can LibreOffice chay headless on Linux.
9. Rui ro loi:
   - Cao voi file DOCX phuc tap do LibreOffice render khac Word.
   - `preview_service.convert_docx_to_pdf` xoa PDF cu cung stem trong output dir.
   - Flow moi bien generate local thanh background job, thay doi hanh vi API `/generate`.
   - Download bi rang buoc boi `preview_seen`; neu iframe bi trinh duyet chan hoac preview route loi, user co the bi ket.
10. Rui ro bao mat:
   - Convert file upload bang LibreOffice co surface xu ly file phuc tap; can gioi han timeout da co 180s.
   - File preview/output nam trong temp/session; can dam bao `sessions/` khong commit, hien da ignored.
11. Migration/cai dat/env: co, cai LibreOffice local hoac set `LIBREOFFICE_BIN`; Docker apt install.
12. Kiem thu:
   - Local khong co LibreOffice: generate phai done voi `preview_error` va van tai DOCX duoc.
   - Local co LibreOffice: iframe load PDF va sau do button DOCX enabled.
   - Cloud: worker upload `preview_object`, web preview download tu GCS.
   - Test DOCX co TOC, bang, anh, font Tieng Viet.
13. Code thua/chua hoan thien:
   - `preview_service.find_libreoffice()` khong tim thu muc untracked hien tai `tools/LibreOfficePortableLibreOfficePortable/...`, chi tim `tools/LibreOfficePortable/...` va `tools/libreoffice/...`.
   - Co 2 logic LibreOffice rieng: `preview_service.py` va `ulnl_tool/core/libreoffice.py`.
14. De xuat: KEEP_AFTER_FIX. Nen giu y tuong preview, nhung phai fix path portable/ignore, xac nhan fallback khi preview fail, va test cloud.
15. Do tin cay: Cao.

### G3. Multi-Excel va Step 6 nang cao cho DOCX builder

1. Muc dich: cho phep upload nhieu Excel, map tung heading voi file/sheet/cot rieng, split bang theo cot, chen heading phu, merge cot, hyperlink noi bo, import/export setup.
2. File chinh:
   - `web_app.py`: Modified, `excel_file` thanh list, `session.json`, `/sessions/{session_id}/excel-files`, build Excel metadata theo file_id.
   - `generation_service.py`: Modified, `normalize_excel_sources`, `resolve_excel_source_path`, advanced selection payload.
   - `templates/configure.html`: Modified rat lon, them controls Step 6 nang cao, import/export setup, add Excel.
   - `docx_template_builder_ver0.9.2.py`: Modified, them logic insertion nang cao trong engine.
3. Truoc thay doi: moi session co 1 Excel; config insert don gian theo sheet/cot/range.
4. Sau thay doi: session co nhieu Excel, moi heading chua `file_id` va advanced config; generate can resolve source phu.
5. Co duoc su dung: co. Route `/configure`, `/generate` va `generation_service.generate_docx_from_payload` da truyen `excel_paths` list/dict.
6. Dependency/cau hinh moi: khong them package, nhung payload client/server thay doi.
7. Anh huong local: session cu khong co `session.json` van fallback tim file; tuy nhien UI/API `/configure` doi `excel_file` thanh multiple list.
8. Anh huong Docker/Cloud: cloud manifest them `excel_files`; worker co fallback manifest cu.
9. Rui ro loi:
   - Cao do surface lon va diff lon.
   - `validate_payload()` goi `build_headings_config(..., {})`; voi heading insert ma khong co excel_sources co the validate khac runtime.
   - Mapping theo `file_id = path.name` co rui ro trung ten file, du da co rename upload duplicate.
   - JS lon trong `configure.html` can test browser; syntax Python parse OK nhung JS chua duoc lint/test.
10. Rui ro bao mat: thap-trung binh; upload nhieu file lam tang I/O va session storage.
11. Migration/cai dat/env: khong.
12. Kiem thu:
   - Upload 1 Excel nhu flow cu.
   - Upload 2 Excel, them Excel sau khi configure, generate heading tu file thu hai.
   - Import/export Step 6 setup roi quay lai edit job.
   - Split table, merge columns, hyperlink noi bo voi DOCX mau.
13. Code thua/chua hoan thien:
   - Duplicate logic `build_headings_config` giua `web_app.py` va `generation_service.py`.
   - `step6_setup_json` duoc dua vao payload nhu `_step6_setup_payload`; can xac nhan co can luu vao output hay chi UI state.
14. De xuat: KEEP_AFTER_FIX. Nen giu neu day la requirement, nhung can tach test regression truoc commit.
15. Do tin cay: Trung binh.

### G4. Cau hinh muc luc TOC va cap nhat field khi mo Word

1. Muc dich: them buoc cau hinh TOC1..TOC9, gioi han TOC vao vung content, danh dau field dirty va update khi mo Word.
2. File chinh:
   - `docx_template_builder_ver0.9.2.py`: Modified, `default_toc_settings`, `_apply_toc_styles`, `_prepare_toc_fields_for_update`, bookmark/content range, numbering suffix spaces.
   - `generation_service.py`: Modified, `apply_toc_settings_config`.
   - `templates/configure.html`: Modified, buoc 7 TOC va buoc 8 Preview.
   - `preview_service.py`: Untracked, uu tien Word COM update page number truoc khi LibreOffice convert.
3. Truoc thay doi: tool 7 buoc, khong co cau hinh TOC rieng.
4. Sau thay doi: tool 8 buoc, TOC style thanh mot phan payload va output DOCX.
5. Co duoc su dung: co. Generate goi `_apply_toc_styles` va `_prepare_toc_fields_for_update`.
6. Dependency/cau hinh moi: tuy chon `DOCX_BUILDER_MARK_TOC_DIRTY`, `DOCX_BUILDER_WORD_TOC_UPDATE`; Microsoft Word tren Windows neu muon update page number chuan.
7. Anh huong local: output DOCX co the yeu cau Word update field khi mo; preview PDF co the phu thuoc Word/LibreOffice.
8. Anh huong Docker/Cloud: Linux khong co Word COM, chi co LibreOffice preview; TOC page number co the khac Word.
9. Rui ro loi:
   - Cao do sua OOXML truc tiep, lien quan bookmark, field instruction, styles, numbering.
   - Co the chen page break truoc content dau tien, lam thay doi pagination.
10. Rui ro bao mat: thap.
11. Migration/cai dat/env: khong bat buoc, nhung preview chuan can Word/LibreOffice.
12. Kiem thu:
   - DOCX co Muc luc, heading intro, heading content.
   - DOCX khong co Muc luc.
   - Heading co auto numbering va TOC cached result.
   - So sanh mo bang Word va PDF preview.
13. Code thua/chua hoan thien: logic TOC nam trong file engine lon, chua co unit test OOXML.
14. De xuat: KEEP_AFTER_FIX.
15. Do tin cay: Trung binh.

### G5. Them tool CSDL/DQC/STD

1. Muc dich: parse tai lieu thiet ke CSDL `.docx`, gan rule DQC/STD, xuat Excel.
2. File chinh:
   - `tk_csdl_tool.py`: Untracked, core parser/rule/writer/CLI.
   - `templates/csdl_tool.html`: Untracked, UI cau hinh.
   - `templates/csdl_job_status.html`: Untracked.
   - `web_app.py`: Modified, route `/csdl`, `/csdl/preview`, `/csdl/generate`, `/csdl/jobs/{job_id}`.
3. Truoc thay doi: khong co module CSDL trong web app.
4. Sau thay doi: CSDL la module rieng trong Solution Hub.
5. Co duoc su dung: co. `web_app.py` dynamic load `tk_csdl_tool.py`.
6. Dependency/cau hinh moi: dung `python-docx`, `openpyxl` da co trong `requirements.txt`.
7. Anh huong local: import `tk_csdl_tool.py` chay `_check_dependencies(verbose=False)` ngay khi load; neu dependency thieu app co the loi luc startup.
8. Anh huong Docker/Cloud: same dependency da co; job CSDL hien local background task trong web process, chua thay cloud-worker rieng.
9. Rui ro loi:
   - Trung binh-cao do parser format A/B lon, can file mau.
   - Encoding terminal hien mojibake khi doc file; chua du thong tin ket luan file hỏng, can xac minh tren editor/browser.
10. Rui ro bao mat: thap-trung binh do parse DOCX upload.
11. Migration/cai dat/env: khong.
12. Kiem thu: upload 1 va nhieu DOCX mau, preview rows, generate overwrite/merge, download Excel, custom rule config/table config.
13. Code thua/chua hoan thien: la file don rat lon, co the can tach module sau khi nghiep vu on dinh.
14. Quyet dinh: NEED_DECISION. Can xac nhan day la module chinh thuc hay prototype.
15. Do tin cay: Trung binh.

### G6. Them tool Data Zone Raw/Work

1. Muc dich: tu Excel nguon tao 2 workbook phu luc Raw Zone va Work Zone.
2. File chinh:
   - `data_zone_builder.py`: Untracked.
   - `templates/data_zone_tool.html`: Untracked.
   - `web_app.py`: Modified, route `/data-zone`, `/data-zone/inspect`, `/data-zone/generate`.
3. Truoc thay doi: khong co module Data Zone.
4. Sau thay doi: user upload Excel, chon sheet/options, generate zip hoac workbook Raw/Work.
5. Co duoc su dung: co. `web_app.py` import cac function tu `data_zone_builder`.
6. Dependency/cau hinh moi: `openpyxl` da co.
7. Anh huong local/Docker: them workflow xu ly Excel; output trong session temp.
8. Rui ro loi: trung binh, mapping header phu thuoc ten cot va tieng Viet; can file mau.
9. Rui ro bao mat: thap-trung binh do upload Excel.
10. Migration/cai dat/env: khong.
11. Kiem thu: file nguon du cot, file thieu cot, sheet nhieu bang, generate Raw/Work, mo Excel output.
12. Code thua/chua hoan thien: chua co test mapping header; output text co the can review nghiep vu.
13. Quyet dinh: NEED_DECISION.
14. Do tin cay: Trung binh.

### G7. Them tool ULNL

1. Muc dich: tao file Uoc luong no luc tu danh sach chuc nang, classify ma CN, build Excel tu template va preview PDF.
2. File chinh:
   - `ulnl_tool/core/data_model.py`, `classifier.py`, `excel_io.py`, `libreoffice.py`: Untracked.
   - `ulnl_tool/scripts/add_functions.py`, `update_phi_cn_params.py`, `update_tong_hop.py`: Untracked.
   - `templates/ulnl_tool.html`: Untracked.
   - `web_app.py`: Modified, route `/ulnl/import`, `/ulnl/classify`, `/ulnl/generate`, `/ulnl/preview-pdf`.
   - `ulnl_tool/resources/ULNL_template_base.xlsx`, `viettel_logo.png`: Added trong local commit L1.
3. Truoc thay doi: khong co module ULNL.
4. Sau thay doi: ULNL la workflow rieng, phu thuoc resource template.
5. Co duoc su dung: co. `web_app.py` import va goi `build_ulnl_file`, `import_functions_from_excel`, `xlsx_to_pdf`.
6. Dependency/cau hinh moi: `openpyxl`; LibreOffice de preview PDF; template Excel resource.
7. Anh huong local: can commit ca code untracked lan resources da local-commit, neu khong module se khong chay o may khac.
8. Anh huong Docker/Cloud: can resources trong Docker context, L1 da cho phep; Docker cai LibreOffice ho tro preview.
9. Rui ro loi:
   - Cao o muc commit planning: resources da nam trong local commit nhung code ULNL van untracked, de push le se lam repo co resource ma khong co feature hoan chinh.
   - Co hai implementation LibreOffice rieng voi path/env khac nhau: `LIBREOFFICE_BIN` va `ULNL_SOFFICE_PATH`.
10. Rui ro bao mat: thap; khong thay secret.
11. Migration/cai dat/env: co the can `ULNL_SOFFICE_PATH` neu preview ULNL local khong tim LibreOffice.
12. Kiem thu: import Excel danh sach chuc nang, classify rules, edit params, generate `.xlsx`, preview PDF.
13. Code thua/chua hoan thien: can hop nhat LibreOffice detection va xac minh encoding UI/text.
14. Quyet dinh: NEED_DECISION. Neu giu Solution Hub thi nen commit tron bo ULNL code + resources; neu khong, bo ca resources da local commit.
15. Do tin cay: Trung binh.

### G8. Cai thien deploy Cloud Run/GCS

1. Muc dich: ho tro LibreOffice tren Cloud Run, preview object tren GCS, va deploy script than thien Windows.
2. File chinh:
   - `Dockerfile`: Modified, apt install LibreOffice/fonts.
   - `scripts/deploy_web.sh`: Modified, doi default `WEB_SERVICE` thanh `solutionhub`, them path `gcloud.cmd`.
   - `storage_backend.py`: Modified, `job_preview_object`.
   - `job_worker.py`: Modified, nhieu Excel + preview PDF.
3. Truoc thay doi: image nhe hon, worker generate DOCX, deploy service default `tool-insert-excel-web`.
4. Sau thay doi: image co LibreOffice, worker upload preview, deploy mac dinh vao `solutionhub`.
5. Co duoc su dung: co, lien quan G2/G3.
6. Dependency/cau hinh moi: apt packages LibreOffice/fonts; Cloud Run memory/disk co the can tang.
7. Anh huong local: Docker build lau/nang hon.
8. Anh huong cloud: service name doi co the tao/deploy sai service neu pipeline dang ky vong ten cu.
9. Rui ro loi: trung binh.
10. Rui ro bao mat: thap.
11. Migration/cai dat/env: co the can cap nhat CI/CD variables `WEB_SERVICE`, memory/timeouts.
12. Kiem thu: `docker build`, run container, hit `/health`, cloud worker job voi DOCX mau.
13. Code thua/chua hoan thien: chua thay update README deploy cho service name moi.
14. De xuat: KEEP_AFTER_FIX. Giu neu da doi product thanh Solution Hub, nhung can xac nhan service name.
15. Do tin cay: Cao.

### G10. LibreOffice portable bung vao repo

1. Muc dich: co ve la copy LibreOffice Portable local de test preview.
2. File chinh:
   - `tools/LibreOfficePortableLibreOfficePortable/**`: Untracked 260 file, tong khoang 4.1 MB.
3. Truoc thay doi: repo khong theo doi bundle nay.
4. Sau thay doi: Git thay 260 file untracked vi pattern ignore hien tai khong khop.
5. Co duoc su dung: chua chac. `preview_service.py` khong tim path nay; README cung huong dan `tools\LibreOfficePortable\...`, khong phai `tools\LibreOfficePortableLibreOfficePortable\...`.
6. Dependency/cau hinh moi: none neu khong commit; neu dung thi can path dung.
7. Anh huong local: co the la artifact giai nen sai thu muc.
8. Anh huong Docker/Cloud: khong nen dua vao image; Docker da cai LibreOffice qua apt.
9. Rui ro loi: cao neu commit vi binary/tool bundle khong can thiet, noisy, co the co license/distribution concern.
10. Rui ro bao mat: chua thay secret ro rang; co file `help.key_` la false positive trong help registry, khong phai private key theo bang chung hien tai.
11. Migration/cai dat/env: khong.
12. Kiem thu: khong can test neu ignore/drop.
13. Code thua/chua hoan thien: pattern `.gitignore` hien co `tools/libreoffice/` va `tools/LibreOfficePortable/` nhung thieu `tools/LibreOfficePortableLibreOfficePortable/`.
14. De xuat: IGNORE_FROM_GIT. Khong commit; them pattern ignore hoac di chuyen ra ngoai repo.
15. Do tin cay: Cao.

## 4. Danh sach file khong nen commit

| File/thu muc | Ly do | Dang duoc .gitignore bo qua chua | De xuat |
|--------------|-------|----------------------------------|---------|
| `tools/LibreOfficePortableLibreOfficePortable/` | Bundle LibreOffice Portable untracked, 260 file binary/config, khong khop path code dang tim | Chua | Them pattern ignore hoac di chuyen ra ngoai repo |
| `.venv/` | Moi truong ao Python | Co | Giu ignore |
| `.venv_old_linux/` | Moi truong ao/backup Python | Co | Giu ignore |
| `__pycache__/` | Python bytecode generated | Co | Giu ignore |
| `ulnl_tool/__pycache__/`, `ulnl_tool/core/__pycache__/`, `ulnl_tool/scripts/__pycache__/` | Python bytecode generated | Co qua `__pycache__/` | Giu ignore |
| `sessions/` | Du lieu upload/output runtime | Co | Giu ignore |
| `[Upload được] ... .docx` | File Word upload test/du lieu nguoi dung | Co qua `*.docx` | Giu ignore |
| `[[Upload được] ... .xlsx` | File Excel upload test/du lieu nguoi dung | Co qua `*.xlsx` | Giu ignore |
| `~$oke_output_portrait_7cols.docx` | File lock/temp cua Word | Co qua `~$*.docx` va `*.docx` | Giu ignore |
| `local_server.err.log`, `local_server.out.log` | Log local | Co qua `local_server.*.log` | Giu ignore |
| `.env` / `.env.*` | Co the chua secret/env local | Co, tru `.env.example` | Giu ignore |
| `*.pyc` | Bytecode generated | Co | Giu ignore |
| `*.zip`, `Archive.zip` | Output/archive | Co | Giu ignore |

Noi dung nen bo sung vao `.gitignore` (chi de xuat, chua tu sua):

```gitignore
# Local LibreOffice portable folders accidentally extracted into repo
tools/LibreOfficePortableLibreOfficePortable/
tools/LibreOfficePortable*/
```

Can can nhac pattern thu hai: `tools/LibreOfficePortable*/` rong hon, se ignore ca `tools/LibreOfficePortable/` va bien the sai ten. Neu sau nay muon commit mot wrapper script trong `tools/LibreOfficePortable...` thi dung pattern cu the hon.

Tuong ung nen bo sung vao `.dockerignore`:

```gitignore
tools/LibreOfficePortableLibreOfficePortable/
tools/LibreOfficePortable*/
```

## 5. Ke hoach xu ly

### Nen giu ngay

- G9 - Tai lieu/UI guideline:
  - `README.md` phan LibreOffice preview.
  - `UI_GUIDELINES.md`, neu team thong nhat dung lam chuan UI.

### Nen sua truoc khi giu

- L1 - Commit local resources/ignore:
  - Bo sung ignore dung cho `tools/LibreOfficePortableLibreOfficePortable/`.
  - Xac minh resources ULNL la file template/logo duoc phep commit.
- G2 - Preview PDF:
  - Hop nhat path/env LibreOffice giua DOCX preview va ULNL preview.
  - Kiem tra fallback download khi preview fail.
  - Test Docker/Cloud Run voi LibreOffice headless.
- G3 - Multi-Excel/Step 6 advanced:
  - Test regression flow 1 Excel cu.
  - Test add Excel sau configure va generate multi-source.
  - Can nhac tach bot duplicate logic `build_headings_config`.
- G4 - TOC:
  - Test voi DOCX co/khong co muc luc, co heading intro/content, co numbering.
  - So sanh output Word voi PDF preview.
- G8 - Cloud deploy:
  - Xac nhan `WEB_SERVICE=solutionhub` co dung target production khong.
  - Cap nhat README/deploy notes neu service name doi.

### Can toi quyet dinh

1. Co doi san pham thanh Solution Hub khong?
   - Phuong an A: Giu Solution Hub. Anh huong: commit G1 + cac module G5/G6/G7, trang `/` la dashboard module.
   - Phuong an B: Giu DOCX Template Builder la trang chinh. Anh huong: bo/doi lai `templates/index.html`, route module moi co the de an hoac tach branch.

2. Co dua CSDL/Data Zone/ULNL vao cung repo/app production khong?
   - Phuong an A: Giu trong app hien tai. Anh huong: can commit tron bo code, template, UI, test tung module.
   - Phuong an B: Tach thanh branch/app rieng. Anh huong: giam rui ro cho tool DOCX hien tai, nhung can tach routing/UI/deploy.

3. Preview PDF co bat buoc truoc download DOCX khong?
   - Phuong an A: Bat buoc. Anh huong: an toan nghiep vu hon nhung phu thuoc LibreOffice/iframe.
   - Phuong an B: Chi la tuy chon. Anh huong: it chan user hon; can sua backend/UI khong enforce `preview_seen`.

4. ULNL resources da local commit co nen push cung code ULNL khong?
   - Phuong an A: Push cung code ULNL. Anh huong: feature day du.
   - Phuong an B: Khong giu ULNL luc nay. Anh huong: can drop/revert local commit L1 hoac tach commit resources sang branch khac.

### Nen loai bo

- G10 - `tools/LibreOfficePortableLibreOfficePortable/**`:
  - Ly do: binary/tool bundle local, 260 file, khong khop path code, khong nen commit.
  - Cach xu ly sau khi ban quyet: ignore hoac xoa khoi workspace bang lenh Git/shell duoc duyet rieng. Audit nay khong xoa.

### Nen dua vao `.gitignore`

Pattern de xuat:

```gitignore
tools/LibreOfficePortableLibreOfficePortable/
tools/LibreOfficePortable*/
```

Pattern hien da co va nen giu:

```gitignore
.venv/
.venv_old_linux/
__pycache__/
*.pyc
sessions/
*.docx
*.xlsx
*.xlsm
!ulnl_tool/resources/*.xlsx
!ulnl_tool/resources/*.xlsm
!ulnl_tool/resources/*.png
local_server.*.log
.env
.env.*
!.env.example
```

## 6. Lenh Git de xuat

Chi de xuat lenh. Khong chay cac lenh stage/drop/clean/restore/reset/commit/push trong audit nay.

Kiem tra lai 4 lop thay doi:

```bash
git log --oneline origin/main..HEAD
git diff --stat origin/main..HEAD
git diff --cached --stat
git diff --stat
git ls-files --others --exclude-standard
```

Stage rieng commit tiep theo cho ignore fix:

```bash
git add .gitignore .dockerignore
```

Stage nhom preview PDF/DOCX job:

```bash
git add preview_service.py env_loader.py .env.example Dockerfile README.md storage_backend.py job_worker.py templates/job_status.html web_app.py
```

Luu y: `web_app.py` cung lien quan nhieu nhom khac, nen neu muon commit tach nho can dung `git add -p web_app.py` va review hunk.

Stage nhom multi-Excel/TOC DOCX builder:

```bash
git add generation_service.py docx_template_builder_ver0.9.2.py templates/configure.html web_app.py
```

Stage nhom Solution Hub UI:

```bash
git add templates/index.html templates/tktkl.html web_app.py UI_GUIDELINES.md
```

Stage nhom CSDL:

```bash
git add tk_csdl_tool.py templates/csdl_tool.html templates/csdl_job_status.html web_app.py
```

Stage nhom Data Zone:

```bash
git add data_zone_builder.py templates/data_zone_tool.html web_app.py
```

Stage nhom ULNL:

```bash
git add ulnl_tool/core ulnl_tool/scripts ulnl_tool/__init__.py templates/ulnl_tool.html web_app.py
```

Xem truoc untracked LibreOffice portable:

```bash
git status --short tools/LibreOfficePortableLibreOfficePortable/
git ls-files --others --exclude-standard tools/LibreOfficePortableLibreOfficePortable/
```

Neu ban quyet dinh loai bo khoi workspace, xem truoc danh sach se xoa bang dry-run truoc:

```bash
git clean -nd tools/LibreOfficePortableLibreOfficePortable/
```

Sau khi da chac chan va co backup/khong can nua, lenh xoa de xuat la:

```bash
git clean -fd tools/LibreOfficePortableLibreOfficePortable/
```

Bo thay doi tung nhom neu quyet dinh drop (de xuat, chua chay):

```bash
git restore templates/index.html
git restore templates/configure.html templates/job_status.html
git restore web_app.py generation_service.py docx_template_builder_ver0.9.2.py
git restore Dockerfile scripts/deploy_web.sh storage_backend.py job_worker.py .env.example README.md
```

Voi file untracked neu muon xoa sau khi xem dry-run:

```bash
git clean -nd UI_GUIDELINES.md data_zone_builder.py env_loader.py preview_service.py tk_csdl_tool.py templates/csdl_job_status.html templates/csdl_tool.html templates/data_zone_tool.html templates/tktkl.html templates/ulnl_tool.html ulnl_tool/
```

Can than: lenh `git clean -fd ulnl_tool/` co the xoa ca file untracked trong `ulnl_tool` nhung khong xoa resources da tracked trong local commit. Neu dung, hay xem dry-run truoc.

## Ket luan ngan

Thay doi hien tai la mot dot mo rong lon tu DOCX Template Builder thanh Solution Hub co 4 module, cong them preview PDF/TOC/multi-Excel. Huong thay doi co gia tri, nhung chua nen commit/push nguyen trang vi:

- Co 260 file LibreOffice Portable untracked khong nen commit.
- Co 1 commit local da them resource ULNL trong khi code ULNL van untracked.
- Nhieu thay doi nghiep vu/UX lon can quyet dinh: Solution Hub, CSDL, Data Zone, ULNL, preview bat buoc.
- Cac phan rui ro cao can test voi file mau thuc te: DOCX TOC, multi-Excel, preview PDF, Cloud Run worker.

