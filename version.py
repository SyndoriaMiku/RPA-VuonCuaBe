"""Phiên bản Tool. `BUILD` được build.ps1 ghi đè tạm thời khi đóng gói."""

__version__ = "1.4.5"
BUILD = "dev"


def window_title(file_title=""):
    """Tiêu đề cửa sổ: kèm mã build khi chạy bản đóng gói, kèm tiêu đề file khi đã chọn file."""
    title = f"Auto Robot v{__version__}"
    if file_title:
        return f"{title} - {file_title}"
    if BUILD != "dev":
        title = f"{title} ({BUILD})"
    return title


def build_text():
    """Dòng phiên bản nhỏ ở góc cửa sổ và trong nhật ký."""
    return f"v{__version__} ({BUILD})"
