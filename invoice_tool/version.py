from __future__ import annotations


__version__ = "6.1.1"
APP_VERSION = f"v{__version__}"
PRODUCT_NAME = "发票处理工具箱"
WINDOWS_EXE_BASENAME = f"invoice-pdf-tool-{APP_VERSION}-windows-x64"
RELEASE_SUMMARY = (
    "v6.1.1 重点更新\n\n"
    "• 修复黑夜主题中原生分组容器残留 Windows 白色默认背景的问题\n"
    "• 统一文件路径、工作簿分析、结果摘要和设置卡片的深色视觉层级\n"
    "• 修复浅色文字叠加白色卡片造成的内容不可读问题\n"
    "• 统一只读下拉框与页面容器的黑夜主题背景\n"
    "• 新增黑夜主题容器颜色、标题对比度和下拉框状态回归测试\n\n"
    "反馈问题时，请从“设置与诊断”导出脱敏诊断包并交给软件提供方。"
)


__all__ = [
    "APP_VERSION",
    "PRODUCT_NAME",
    "RELEASE_SUMMARY",
    "WINDOWS_EXE_BASENAME",
    "__version__",
]
