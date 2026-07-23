from __future__ import annotations


__version__ = "6.0.0"
APP_VERSION = f"v{__version__}"
PRODUCT_NAME = "发票处理工具箱"
WINDOWS_EXE_BASENAME = f"invoice-pdf-tool-{APP_VERSION}-windows-x64"
RELEASE_SUMMARY = (
    "v6.0.0 重点更新\n\n"
    "• 文件移动、复制和回滚增加受信目录与内容指纹保护\n"
    "• 后台预览、暂停/取消、失败重试和异常中断恢复\n"
    "• Excel 筛选执行前强制当前预览与最终确认\n"
    "• 新版任务工作台、历史详情、配置导入导出和脱敏诊断\n"
    "• 大目录与大型 Excel 性能、.xls 支持及 Windows 打包可靠性改进\n\n"
    "反馈问题时，请从“设置与诊断”导出脱敏诊断包并交给软件提供方。"
)


__all__ = [
    "APP_VERSION",
    "PRODUCT_NAME",
    "RELEASE_SUMMARY",
    "WINDOWS_EXE_BASENAME",
    "__version__",
]
