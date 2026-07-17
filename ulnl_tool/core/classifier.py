"""
core/classifier.py - Rule-based engine để gán mã chức năng (ma_cn) tự động.

Logic: mỗi rule có:
- priority: số thứ tự ưu tiên (rule có priority thấp hơn xử lý trước)
- keywords: dict các nhóm keyword (any/all/none) - kiểm tra trên name + desc
- conditions: dict điều kiện bổ sung (regex, platform, group_code, ...)
- result: mã CN trả về
- comment: ghi chú cho user

Rule structure (JSON):
{
  "id": "auth-login",
  "priority": 100,
  "name": "Đăng nhập / mật khẩu",
  "enabled": true,
  "keywords": {
    "any_of": ["đăng nhập", "đăng xuất", "đổi mật khẩu", "quên mật khẩu",
               "login", "logout", "change password", "forgot password"],
    "all_of": [],
    "none_of": []
  },
  "name_starts_with": [],
  "name_equals": [],
  "platform": "web|mobile|both",
  "result": "CN_WEB2",
  "comment": "Auth functions thường có form đơn giản → CN_WEB2"
}

Áp dụng theo thứ tự priority tăng dần, rule nào match đầu tiên thắng.
Có rule "default" cuối cùng (priority 9999) trả CN_WEB2 / CN_MOBILE Client2.
"""
from __future__ import annotations
import json
import re
from dataclasses import dataclass, field, asdict
from typing import Optional
from pathlib import Path


@dataclass
class ClassificationRule:
    id: str
    name: str
    result: str
    priority: int = 100
    enabled: bool = True
    keywords_any: list = field(default_factory=list)
    keywords_all: list = field(default_factory=list)
    keywords_none: list = field(default_factory=list)
    name_starts_with: list = field(default_factory=list)
    name_equals: list = field(default_factory=list)
    name_regex: Optional[str] = None
    platform: str = "both"
    group_code: Optional[str] = None
    comment: str = ""

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "ClassificationRule":
        return cls(**{k: v for k, v in d.items() if k in cls.__annotations__})

    def matches(self, name: str, desc: str, platform: str, group_code: str) -> bool:
        if not self.enabled:
            return False
        if self.platform != "both" and self.platform != platform:
            return False
        if self.group_code and self.group_code != group_code:
            return False

        n_lower = (name or "").lower().strip()
        text = (n_lower + " " + (desc or "").lower())

        def kw_match(kw: str, text: str) -> bool:
            """Check keyword match. Với kw ngắn (<5 ký tự, không space, ASCII)
            dùng word-boundary để tránh substring false-positive
            (vd: 'word' không match 'keyword'). Với kw dài hoặc có space/dấu, dùng substring."""
            k = kw.lower()
            if len(k) < 5 and " " not in k and k.isascii() and k.isalnum():
                return re.search(r"\b" + re.escape(k) + r"\b", text) is not None
            return k in text

        # Logic: name_equals / name_starts_with / name_regex / keywords là các cách
        # match khác nhau. Nếu rule có nhiều cách, MỘT trong số chúng match đủ.
        has_match_criterion = False
        criterion_matched = False

        if self.name_equals:
            has_match_criterion = True
            if any(s.lower() == n_lower for s in self.name_equals):
                criterion_matched = True

        if self.name_starts_with:
            has_match_criterion = True
            if any(n_lower.startswith(s.lower()) for s in self.name_starts_with):
                criterion_matched = True

        if self.name_regex:
            has_match_criterion = True
            try:
                if re.search(self.name_regex, name or "", re.IGNORECASE):
                    criterion_matched = True
            except re.error:
                pass

        if self.keywords_any:
            has_match_criterion = True
            if any(kw_match(kw, text) for kw in self.keywords_any):
                criterion_matched = True

        if has_match_criterion and not criterion_matched:
            return False

        # Filters
        if self.keywords_all:
            if not all(kw_match(kw, text) for kw in self.keywords_all):
                return False

        if self.keywords_none:
            if any(kw_match(kw, text) for kw in self.keywords_none):
                return False

        return True


# ===== Default rule set =====
# Đã chuyển toàn bộ heuristic từ Python sang declarative rules.
# Mỗi rule có keywords Việt + Anh để hỗ trợ cả 2 ngôn ngữ.
DEFAULT_RULES: list[ClassificationRule] = [
    # ===== Dashboard (priority cao nhất - dễ nhầm với báo cáo) =====
    ClassificationRule(
        id="dashboard-exec",
        name="Dashboard điều hành/lãnh đạo",
        priority=10,
        keywords_any=["dashboard tổng quan", "dashboard điều hành", "dashboard lãnh đạo",
                      "executive dashboard", "leadership dashboard"],
        result="DV_TABLEAU_DASHBOARD2",
        comment="Dashboard cho lãnh đạo - phức tạp"
    ),
    ClassificationRule(
        id="dashboard-general",
        name="Dashboard chung",
        priority=11,
        keywords_any=["dashboard"],
        result="DV_TABLEAU_DASHBOARD1",
        comment="Dashboard thông thường"
    ),

    # ===== Báo cáo =====
    ClassificationRule(
        id="report-complex",
        name="Báo cáo phức tạp (tổng hợp/phân tích)",
        priority=20,
        keywords_any=["báo cáo tổng hợp", "báo cáo phân tích", "báo cáo đa chiều",
                      "summary report", "analysis report", "drill-down report",
                      "báo cáo lãnh đạo", "leadership report"],
        result="BC3",
        comment="Báo cáo phức tạp với drill-down, pivot, chart"
    ),
    ClassificationRule(
        id="report-standard",
        name="Báo cáo thông thường",
        priority=21,
        keywords_any=["báo cáo", "biểu mẫu báo cáo", "kết quả thẩm định", "thống kê",
                      "report", "statistics"],
        keywords_none=["workflow trình", "phê duyệt báo cáo"],
        result="BC2",
        comment="Báo cáo trung bình"
    ),

    # ===== Import / Export =====
    ClassificationRule(
        id="import-excel",
        name="Import file Excel",
        priority=30,
        keywords_all=["import"],
        keywords_any=["excel", "xlsx", "csv"],
        result="DP_IMP2",
        comment="Import dữ liệu từ Excel - data pipeline"
    ),
    ClassificationRule(
        id="export-excel",
        name="Export file Excel",
        priority=31,
        keywords_all=["export"],
        keywords_any=["excel", "xlsx", "csv"],
        result="DP_EXP1",
        comment="Export ra Excel"
    ),
    ClassificationRule(
        id="export-pdf-word",
        name="Export file PDF/Word",
        priority=32,
        name_starts_with=["export", "xuất file"],
        keywords_any=["pdf", "word", "docx"],
        result="BC2",
        comment="Export PDF/Word - bản chất là báo cáo"
    ),

    # ===== Tổng hợp data =====
    ClassificationRule(
        id="aggregate",
        name="Tổng hợp / rollup dữ liệu",
        priority=40,
        keywords_any=["tổng hợp dữ liệu", "tổng hợp theo", "aggregate", "rollup",
                      "cộng dồn", "tích lũy"],
        result="DP_AGG2",
        comment="Pipeline tổng hợp data"
    ),

    # ===== Form Designer =====
    ClassificationRule(
        id="form-designer",
        name="Form Designer / cấu hình công thức",
        priority=50,
        keywords_any=["form designer", "cấu trúc biểu mẫu", "cấu trúc cột",
                      "cấu hình cấu trúc", "cấu hình công thức",
                      "form builder", "formula configuration"],
        result="CN_WEB3",
        comment="Form designer là chức năng phức tạp"
    ),

    # ===== Workflow phê duyệt =====
    ClassificationRule(
        id="workflow-approval",
        name="Workflow phê duyệt phức tạp",
        priority=60,
        keywords_any=["workflow phê duyệt", "phê duyệt nhiều cấp", "phê duyệt cuối",
                      "trình hđqt", "trình tgđ", "hội đồng thẩm định",
                      "multi-level approval", "approval workflow"],
        result="NVJ2-PTM",
        comment="Workflow phê duyệt - logic nghiệp vụ phức tạp"
    ),

    # ===== Thông báo / Notification engine =====
    ClassificationRule(
        id="notification-send",
        name="Gửi thông báo (email/SMS/push)",
        priority=70,
        name_starts_with=["thông báo", "gửi mail", "gửi sms", "send notification",
                          "send email", "send sms"],
        keywords_any=["push notification"],
        keywords_none=["cấu hình", "xem", "list", "danh sách"],
        result="NVJ2-PTM",
        comment="Notification engine - logic backend"
    ),
    ClassificationRule(
        id="notification-channel",
        name="Thông báo qua email/mobile push",
        priority=71,
        keywords_any=["qua email", "qua mobile push", "via email", "via push"],
        result="NVJ2-PTM",
    ),

    # ===== Tích hợp hệ thống ngoài =====
    ClassificationRule(
        id="integration-config",
        name="Cấu hình tích hợp (form)",
        priority=80,
        keywords_all=["cấu hình"],
        keywords_any=["e-office", "tích hợp", "erp", "api", "callback", "webhook",
                      "integration", "endpoint"],
        result="CN_WEB2",
        comment="Cấu hình tích hợp - form web"
    ),
    ClassificationRule(
        id="integration-call",
        name="Gọi API / tải file qua API",
        priority=81,
        keywords_any=["gọi api", "lấy file", "tải file", "call api", "fetch file",
                      "download file"],
        keywords_all=["api"],
        result="NVJ2-PTM",
        comment="Gọi API ngoài"
    ),
    ClassificationRule(
        id="integration-complex",
        name="Tích hợp hệ thống ngoài (logic phức tạp)",
        priority=82,
        keywords_any=["e-office", "tích hợp", "erp", "callback", "webhook"],
        result="NVJ3-PTM",
        comment="Tích hợp với hệ thống ngoài, logic phức tạp"
    ),

    # ===== Tính toán / Engine =====
    ClassificationRule(
        id="calc-engine",
        name="Tính toán / Engine",
        priority=90,
        keywords_any=["tính toán", "tự động tính", "rule engine", "engine", "calculation"],
        result="NVJ2-PTM",
    ),

    # ===== Validation =====
    ClassificationRule(
        id="validation",
        name="Kiểm tra / Validation",
        priority=100,
        keywords_any=["kiểm tra logic", "kiểm tra hợp lệ", "validation",
                      "cảnh báo tự động", "cảnh báo sai lệch",
                      "validate", "auto warning"],
        result="NVJ2-PTM",
    ),

    # ===== Auth =====
    ClassificationRule(
        id="auth",
        name="Authentication (đăng nhập/mật khẩu)",
        priority=110,
        name_equals=["đăng nhập", "đăng xuất", "login", "logout", "sign in", "sign out"],
        keywords_any=["đăng nhập", "đăng xuất", "login", "logout",
                      "đổi mật khẩu", "quên mật khẩu", "change password",
                      "forgot password", "reset password"],
        result="CN_WEB2",
    ),

    # ===== CRUD operations =====
    ClassificationRule(
        id="crud-create",
        name="Thêm mới (CRUD - Create)",
        priority=120,
        name_starts_with=["thêm mới", "thêm ", "add ", "create ", "new "],
        result="CN_WEB2",
    ),
    ClassificationRule(
        id="crud-update",
        name="Sửa / Chỉnh sửa (CRUD - Update)",
        priority=121,
        name_starts_with=["sửa ", "chỉnh sửa", "cập nhật", "edit ", "update ", "modify "],
        result="CN_WEB2",
    ),
    ClassificationRule(
        id="crud-delete",
        name="Xóa (CRUD - Delete)",
        priority=122,
        name_starts_with=["xóa ", "delete ", "remove "],
        result="CN_WEB1",
    ),
    ClassificationRule(
        id="lock-unlock",
        name="Khóa / Mở khóa / Kích hoạt",
        priority=123,
        name_starts_with=["khóa", "lock"],
        keywords_any=["vô hiệu hóa", "kích hoạt", "enable", "disable", "activate"],
        result="CN_WEB2",
    ),

    # ===== Xem / Hiển thị / Tra cứu =====
    ClassificationRule(
        id="view-display",
        name="Xem / Hiển thị / Tra cứu",
        priority=130,
        name_starts_with=["xem ", "hiển thị", "tra cứu", "view ", "display ", "show "],
        keywords_any=["tra cứu", "lookup"],
        result="CN_WEB1",
    ),
    ClassificationRule(
        id="search-filter",
        name="Tìm kiếm / Lọc",
        priority=131,
        name_starts_with=["tìm kiếm", "lọc theo", "search", "filter"],
        result="CN_WEB1",
    ),

    # ===== Versioning =====
    ClassificationRule(
        id="version-view",
        name="Xem / So sánh phiên bản",
        priority=140,
        keywords_all=["phiên bản"],
        keywords_any=["xem", "so sánh", "view", "compare"],
        result="CN_WEB1",
    ),
    ClassificationRule(
        id="version-engine",
        name="Versioning engine",
        priority=141,
        keywords_any=["version", "phiên bản", "versioning"],
        result="NVJ1-PTM",
    ),

    # ===== Log / Audit =====
    ClassificationRule(
        id="log-audit",
        name="Ghi log / Audit",
        priority=150,
        keywords_any=["ghi log", "audit", "logging"],
        result="NVJ1-PTM",
    ),
    ClassificationRule(
        id="log-view",
        name="Xem / Tra cứu nhật ký",
        priority=151,
        keywords_all=["nhật ký"],
        keywords_any=["xem", "tra cứu", "tìm", "xuất", "view", "search", "export"],
        result="CN_WEB1",
    ),

    # ===== Phân quyền =====
    ClassificationRule(
        id="permission",
        name="Phân quyền / Gán vai trò",
        priority=160,
        keywords_any=["phân quyền", "gán quyền", "phân vai", "gán vai trò",
                      "permission", "role assignment"],
        result="CN_WEB2",
    ),

    # ===== Cấu hình chung =====
    ClassificationRule(
        id="config",
        name="Cấu hình / Thiết lập / Khai báo",
        priority=170,
        keywords_any=["cấu hình", "thiết lập", "khai báo",
                      "configuration", "configure", "setup"],
        result="CN_WEB2",
    ),

    # ===== Quản lý danh mục =====
    ClassificationRule(
        id="dictionary",
        name="Quản lý danh mục",
        priority=180,
        name_starts_with=["quản lý danh mục", "quản lý loại",
                          "manage category", "manage dictionary"],
        result="CN_WEB2",
    ),

    # ===== Nhắc việc / Reminder =====
    ClassificationRule(
        id="reminder",
        name="Nhắc nhở / Reminder",
        priority=190,
        keywords_any=["nhắc nhở", "nhắc việc", "reminder"],
        result="NVJ2-PTM",
    ),

    # ===== Phê duyệt đơn giản =====
    ClassificationRule(
        id="approve-simple",
        name="Phê duyệt (form đơn)",
        priority=200,
        name_starts_with=["phê duyệt", "approve "],
        result="CN_WEB2",
    ),

    # ===== Tạo / Lập =====
    ClassificationRule(
        id="create-form",
        name="Tạo / Lập (form)",
        priority=210,
        name_starts_with=["tạo ", "lập ", "create ", "establish "],
        result="CN_WEB2",
    ),

    # ===== Submit / Gửi =====
    ClassificationRule(
        id="submit",
        name="Gửi / Nộp / Submit",
        priority=220,
        name_starts_with=["gửi ", "nộp ", "submit", "send"],
        result="CN_WEB2",
    ),

    # ===== Theo dõi trạng thái =====
    ClassificationRule(
        id="track-status",
        name="Theo dõi trạng thái",
        priority=230,
        name_starts_with=["theo dõi", "track", "monitor"],
        keywords_any=["trạng thái", "status"],
        result="CN_WEB1",
    ),

    # ===== MOBILE rules - priority thấp (chạy trước) khi platform=mobile =====
    ClassificationRule(
        id="mobile-server-sync",
        name="Mobile - Sync / Push / API server",
        priority=5,
        platform="mobile",
        keywords_any=["đồng bộ", "sync", "api server", "push notification",
                      "webhook", "callback"],
        result="CN_MOBILE Server2",
    ),
    ClassificationRule(
        id="mobile-server-basic",
        name="Mobile - API đơn giản",
        priority=6,
        platform="mobile",
        keywords_any=["push notification", "api"],
        result="CN_MOBILE Server1",
    ),
    ClassificationRule(
        id="mobile-field-data",
        name="Mobile - Hiện trường (camera/GPS/offline)",
        priority=7,
        platform="mobile",
        keywords_any=["chụp ảnh", "gps", "đính kèm", "minh chứng", "offline",
                      "camera", "geolocation", "attachment"],
        result="CN_MOBILE Client2",
    ),
    ClassificationRule(
        id="mobile-form",
        name="Mobile - Form nhập liệu",
        priority=8,
        platform="mobile",
        keywords_any=["form", "nhập", "cập nhật", "input", "update"],
        keywords_none=["xem", "view"],
        result="CN_MOBILE Client2",
    ),
    ClassificationRule(
        id="mobile-report",
        name="Mobile - Xem báo cáo",
        priority=9,
        platform="mobile",
        keywords_any=["báo cáo", "thống kê", "biểu đồ", "report", "chart"],
        result="CN_MOBILE Client2",
    ),
    ClassificationRule(
        id="mobile-view",
        name="Mobile - Xem / Tra cứu đơn giản",
        priority=300,
        platform="mobile",
        keywords_any=["hiển thị", "xem", "tra cứu", "view", "display"],
        result="CN_MOBILE Client1",
    ),
    ClassificationRule(
        id="mobile-default",
        name="Mobile - Default",
        priority=9990,
        platform="mobile",
        result="CN_MOBILE Client1",
    ),

    # ===== Default cuối cùng cho Web =====
    ClassificationRule(
        id="default-web",
        name="Default (Web - trung bình)",
        priority=9999,
        platform="both",
        result="CN_WEB2",
        comment="Mặc định nếu không rule nào match"
    ),
]


class Classifier:
    """Apply rules theo priority để phân loại mã CN."""

    def __init__(self, rules: Optional[list] = None):
        self.rules = rules if rules is not None else [
            ClassificationRule(**r.to_dict()) for r in DEFAULT_RULES
        ]
        self._sort()

    def _sort(self):
        self.rules.sort(key=lambda r: r.priority)

    def classify(self, name: str, desc: str = "", platform: str = "web",
                 group_code: str = "") -> tuple[str, str]:
        """Trả về (ma_cn, rule_id_matched). Nếu không rule nào match → fallback."""
        for rule in self.rules:
            if rule.matches(name, desc, platform, group_code):
                return rule.result, rule.id
        return ("CN_MOBILE Client1" if platform == "mobile" else "CN_WEB2", "fallback")

    def to_dict(self) -> dict:
        return {"rules": [r.to_dict() for r in self.rules]}

    def to_json(self, path: str | Path):
        Path(path).write_text(json.dumps(self.to_dict(), ensure_ascii=False, indent=2),
                              encoding="utf-8")

    @classmethod
    def from_json(cls, path: str | Path) -> "Classifier":
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        rules = [ClassificationRule.from_dict(r) for r in data["rules"]]
        return cls(rules)

    def add_rule(self, rule: ClassificationRule):
        self.rules.append(rule)
        self._sort()

    def remove_rule(self, rule_id: str) -> bool:
        for i, r in enumerate(self.rules):
            if r.id == rule_id:
                del self.rules[i]
                return True
        return False

    def update_rule(self, rule: ClassificationRule) -> bool:
        for i, r in enumerate(self.rules):
            if r.id == rule.id:
                self.rules[i] = rule
                self._sort()
                return True
        return False
