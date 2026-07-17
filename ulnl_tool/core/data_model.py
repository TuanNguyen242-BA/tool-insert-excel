"""
core/data_model.py - Data models cho ULNL project.

Cấu trúc:
- Function: 1 chức năng (input từ Excel, có cấp 1/2/3, mã CN, mô tả...)
- Project: chứa danh sách Function + settings các bước

Function được tổ chức theo CẤP (1=group, 2=subgroup, 3=item). Khi build ULNL,
hierarchy được suy ra từ thứ tự + cấp (như cách Excel danh sách CN sắp xếp).
"""
from __future__ import annotations
from dataclasses import dataclass, field, asdict
from typing import Optional, List
import json
import copy
from pathlib import Path


@dataclass
class Function:
    """1 dòng chức năng trong danh sách."""
    stt: int = 0
    ma: str = ""              # F1, F1.1, F1.1.1
    cap: int = 3              # 1, 2, 3
    name: str = ""
    desc: str = ""
    # Sau khi enrich:
    ma_cn_web: str = ""       # Mã CN cho phiên bản Web (nếu có)
    ma_cn_mobile: str = ""    # Mã CN cho phiên bản Mobile (nếu có)
    include_web: bool = True
    include_mobile: bool = False
    # % tái sử dụng
    o: float = 0.0    # GP
    p: float = 0.0    # PT
    q: float = 0.0    # KT
    # Thuê ngoài
    tn_gp: str = "Có"
    tn_pt: str = "Có"
    tn_kt: str = "Có"
    # Ghi chú
    tsd_tu: str = ""
    ghi_chu: str = ""

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "Function":
        return cls(**{k: v for k, v in d.items() if k in cls.__annotations__})


@dataclass
class CoverPageInfo:
    """Thông tin Trang bia."""
    project_name: str = "TÊN DỰ ÁN / TÀI LIỆU ULNL"
    project_code: str = "DA-001"
    doc_code: str = "VTNet/PL02/ULNL/01"
    version: str = "Ef"
    nguoi_lap_ten: str = ""
    nguoi_lap_chucvu: str = ""
    nguoi_lap_ngay: str = ""
    nguoi_ktra_ten: str = ""
    nguoi_ktra_chucvu: str = ""
    nguoi_ktra_ngay: str = ""
    nguoi_pduyet_ten: str = ""
    nguoi_pduyet_chucvu: str = ""
    nguoi_pduyet_ngay: str = ""

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "CoverPageInfo":
        return cls(**{k: v for k, v in d.items() if k in cls.__annotations__})


@dataclass
class PhiCNParams:
    """Tham số NL Phi CN (số module, số hệ thống, năm BH...)."""
    n_module: int = 6
    n_upcode: int = 10
    n_bugs: int = 18
    n_year: int = 3
    n_hrs_bh_month: int = 60
    n_sys: int = 2
    pct_qtda: float = 0.10
    # User-defined extra params (cho mở rộng):
    custom_params: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "PhiCNParams":
        return cls(**{k: v for k, v in d.items() if k in cls.__annotations__})


@dataclass
class CellStyle:
    """Style cho 1 vùng cell."""
    font_name: str = "Times New Roman"
    font_size: int = 11
    bold: bool = False
    italic: bool = False
    font_color: str = "000000"          # hex RRGGBB
    bg_color: str = "FFFFFF"            # hex RRGGBB
    horizontal: str = "left"            # left/center/right
    vertical: str = "center"            # top/center/bottom
    wrap_text: bool = True
    border: bool = True

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "CellStyle":
        return cls(**{k: v for k, v in d.items() if k in cls.__annotations__})


@dataclass
class SheetFormatting:
    """Format cho 1 sheet cụ thể."""
    title: CellStyle = field(default_factory=lambda: CellStyle(
        font_size=14, bold=True, horizontal="center", bg_color="FFFFFF", border=False
    ))
    header: CellStyle = field(default_factory=lambda: CellStyle(
        font_size=11, bold=True, italic=True, horizontal="center",
        bg_color="CCFFCC"
    ))
    data: CellStyle = field(default_factory=lambda: CellStyle(
        font_size=10, horizontal="left"
    ))
    total: CellStyle = field(default_factory=lambda: CellStyle(
        font_size=11, bold=True, horizontal="center", bg_color="66CC66"
    ))

    def to_dict(self) -> dict:
        return {
            "title": self.title.to_dict(),
            "header": self.header.to_dict(),
            "data": self.data.to_dict(),
            "total": self.total.to_dict(),
        }

    @classmethod
    def from_dict(cls, d: dict) -> "SheetFormatting":
        return cls(
            title=CellStyle.from_dict(d.get("title", {})),
            header=CellStyle.from_dict(d.get("header", {})),
            data=CellStyle.from_dict(d.get("data", {})),
            total=CellStyle.from_dict(d.get("total", {})),
        )


@dataclass
class FormattingConfig:
    """Toàn bộ formatting cho 4 sheet chính."""
    trang_bia: SheetFormatting = field(default_factory=SheetFormatting)
    tong_hop: SheetFormatting = field(default_factory=SheetFormatting)
    nl_chuc_nang: SheetFormatting = field(default_factory=SheetFormatting)
    nl_phi_cn: SheetFormatting = field(default_factory=SheetFormatting)

    def to_dict(self) -> dict:
        return {
            "trang_bia": self.trang_bia.to_dict(),
            "tong_hop": self.tong_hop.to_dict(),
            "nl_chuc_nang": self.nl_chuc_nang.to_dict(),
            "nl_phi_cn": self.nl_phi_cn.to_dict(),
        }

    @classmethod
    def from_dict(cls, d: dict) -> "FormattingConfig":
        return cls(
            trang_bia=SheetFormatting.from_dict(d.get("trang_bia", {})),
            tong_hop=SheetFormatting.from_dict(d.get("tong_hop", {})),
            nl_chuc_nang=SheetFormatting.from_dict(d.get("nl_chuc_nang", {})),
            nl_phi_cn=SheetFormatting.from_dict(d.get("nl_phi_cn", {})),
        )


@dataclass
class Project:
    """Toàn bộ state của 1 ULNL project."""
    functions: List[Function] = field(default_factory=list)
    cover: CoverPageInfo = field(default_factory=CoverPageInfo)
    phi_cn: PhiCNParams = field(default_factory=PhiCNParams)
    formatting: FormattingConfig = field(default_factory=FormattingConfig)
    enable_web: bool = True
    enable_mobile: bool = False
    classifier_rules: list = field(default_factory=list)   # rules JSON dump

    def to_dict(self) -> dict:
        return {
            "functions": [f.to_dict() for f in self.functions],
            "cover": self.cover.to_dict(),
            "phi_cn": self.phi_cn.to_dict(),
            "formatting": self.formatting.to_dict(),
            "enable_web": self.enable_web,
            "enable_mobile": self.enable_mobile,
            "classifier_rules": self.classifier_rules,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "Project":
        return cls(
            functions=[Function.from_dict(f) for f in d.get("functions", [])],
            cover=CoverPageInfo.from_dict(d.get("cover", {})),
            phi_cn=PhiCNParams.from_dict(d.get("phi_cn", {})),
            formatting=FormattingConfig.from_dict(d.get("formatting", {})),
            enable_web=d.get("enable_web", True),
            enable_mobile=d.get("enable_mobile", False),
            classifier_rules=d.get("classifier_rules", []),
        )

    def save(self, path: str | Path):
        Path(path).write_text(
            json.dumps(self.to_dict(), ensure_ascii=False, indent=2),
            encoding="utf-8"
        )

    @classmethod
    def load(cls, path: str | Path) -> "Project":
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        return cls.from_dict(data)

    def clone(self) -> "Project":
        return copy.deepcopy(self)
