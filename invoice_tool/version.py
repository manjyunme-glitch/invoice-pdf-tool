from __future__ import annotations


__version__ = "6.1.0"
APP_VERSION = f"v{__version__}"
PRODUCT_NAME = "发票处理工具箱"
WINDOWS_EXE_BASENAME = f"invoice-pdf-tool-{APP_VERSION}-windows-x64"
RELEASE_SUMMARY = (
    "v6.1.0 重点更新\n\n"
    "• 修复输入、规则、预览、执行和结果步骤导航的滚动偏移\n"
    "• 修复 ttkbootstrap 覆盖应用配色导致侧栏与内容区域割裂的问题\n"
    "• 提升普通、悬停和禁用按钮在白天与黑夜主题下的文字对比度\n"
    "• 恢复主操作按钮字号层级，并优化页面说明与流程文字排版\n"
    "• 新增步骤定位、主题配色、按钮可读性和字号层级回归测试\n\n"
    "反馈问题时，请从“设置与诊断”导出脱敏诊断包并交给软件提供方。"
)


__all__ = [
    "APP_VERSION",
    "PRODUCT_NAME",
    "RELEASE_SUMMARY",
    "WINDOWS_EXE_BASENAME",
    "__version__",
]
