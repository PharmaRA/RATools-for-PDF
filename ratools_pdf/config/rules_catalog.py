"""处理规则目录：模块分组、规则标题/描述、eCTD 预设。

此前同一批规则的标题在 main_window.MODULES_DATA 与 precheck.PRECHECK_OPTION_TITLES
各维护一份，且已经发生文案漂移（如"书签动作改为新窗口打开" vs "书签动作：新窗口打开"）。
这里是唯一事实来源：UI 勾选框、预检日志、预检建议标题都从本目录读取。
标题以 UI 显示文案为准，保证日志里出现的规则名与界面一致。

本模块禁止依赖 Qt / fitz。
"""

# 每个模块：icon、title、options（id/title/desc 列表）。
# "文档检测"模块没有勾选规则，由专用按钮触发只读扫描。
MODULES = [
    {
        "icon": "👀",
        "title": "初始视图与文档属性",
        "options": [
            {"id": "open_page_first", "title": "设为首页打开", "desc": "强制文档打开时默认显示第一页"},
            {"id": "page_layout_default", "title": "重置页面布局", "desc": "将页面布局恢复为默认"},
            {"id": "zoom_default", "title": "重置缩放比例", "desc": "将打开时的缩放比例设置为默认"},
            {"id": "initial_view_bookmarks_and_page", "title": "设置导览标签", "desc": "包含书签的文档，导览标签设置为书签面板和页面；不包含书签的文档，导览标签设置为页面。"},
            {"id": "collapse_all_bookmarks", "title": "折叠所有书签", "desc": "将书签树默认设置为折叠状态，保持界面整洁"},
            {"id": "title_from_filename", "title": "同步文件名为标题", "desc": "自动将当前PDF的文件名写入文档属性的“标题”元数据中"},
        ],
    },
    {
        "icon": "📄",
        "title": "页面与字体标准化",
        "options": [
            {"id": "page_size_a4", "title": "适配到A4尺寸", "desc": "无法通过预检精确判断是否需要处理；智能模式下若勾选仍会执行"},
            {"id": "page_size_letter", "title": "适配到Letter尺寸", "desc": "按原页面方向等比缩放并居中留白，适配到Letter (信纸) 尺寸，尽量保留全部内容"},
        ],
    },
    {
        "icon": "🔖",
        "title": "书签管理",
        "options": [
            {"id": "bookmark_inherit_zoom", "title": "书签设为承前缩放", "desc": "点击书签跳转时，保持当前页面的缩放比例不变 (Inherit Zoom)"},
            {"id": "bookmark_open_new_window", "title": "书签动作：新窗口打开", "desc": "配置书签的链接跳转默认在新的PDF浏览器窗口中打开"},
            {"id": "bookmark_remove_external_links", "title": "清理书签外部链接", "desc": "移除书签中指向网页或外部文件的URI动作"},
            {"id": "bookmark_remove_invalid", "title": "清理失效书签", "desc": "自动检测并删除未指向任何有效页面或动作的空书签"},
            {"id": "bookmark_remove_unknown_actions", "title": "清理非标准动作书签", "desc": "仅保留内部跳转、外部文档和调用命令，删除其它未知动作"},
        ],
    },
    {
        "icon": "🔗",
        "title": "超链接处理",
        "options": [
            {"id": "link_abs_to_rel_path", "title": "绝对路径转相对路径", "desc": "将外部文件链接的绝对路径自动转换为相对路径"},
            {"id": "link_inherit_zoom", "title": "超链接设为承前缩放", "desc": "点击链接跳转时，保持当前屏幕的视图缩放比例 (Inherit Zoom)"},
            {"id": "link_open_new_window", "title": "链接动作：新窗口打开", "desc": "强制外部文档或网页链接在独立的新窗口中打开"},
            {"id": "link_text_blue", "title": "链接文本设为蓝色", "desc": "自动识别超链接区域并将其文本颜色变更为标准蓝色"},
            {"id": "link_black_border", "title": "链接区域加黑框", "desc": "为所有的有效超链接区域添加1px的黑色实线边框"},
            {"id": "link_bordered_to_blue_border", "title": "标准化有框链接", "desc": "若超链接已存在边框，则统一转为蓝框黑字样式"},
            {"id": "link_unbordered_blue_to_blue_border", "title": "标准化无框蓝字链接", "desc": "若超链接无边框且文字为蓝色，则统一转为蓝框黑字样式"},
            {"id": "link_remove_border", "title": "清除所有链接边框", "desc": "移除文档内所有超链接的可见边框，保持页面排版干净"},
        ],
    },
    {
        "icon": "🛡️",
        "title": "内容合规与安全性",
        "options": [
            {"id": "cleanup_remove_external_uri", "title": "删除外部URI链接", "desc": "清理指向外部网站、邮箱等所有URI类型的超链接"},
            {"id": "cleanup_remove_external_uri_and_text_black", "title": "删除外部URI链接并去色", "desc": "清理URI链接的同时，将该链接对应的文本颜色重置为黑色"},
            {"id": "cleanup_remove_invalid_links", "title": "清理失效超链接", "desc": "自动扫描并移除所有未分配有效动作 (Action) 的空链接"},
            {"id": "cleanup_remove_invalid_links_and_text_black", "title": "清理失效链接并去色", "desc": "移除空链接，并将该区域相关的文本颜色恢复为普通黑色"},
            {"id": "cleanup_remove_unknown_action_links", "title": "清理非标准动作链接", "desc": "仅保留内部/外部跳转和执行动作，移除其它所有的特殊行为"},
            {"id": "cleanup_remove_dynamic_content", "title": "彻底清除动态内容 (JS/3D)", "desc": "删除文档内所有的JavaScript脚本、3D模型等交互元素以满足安全合规"},
            {"id": "cleanup_remove_attachments", "title": "移除所有内嵌附件", "desc": "清理PDF内部打包的所有附加文件 (.zip, .xml 等)"},
            {"id": "cleanup_remove_tags", "title": "移除结构化标签", "desc": "删除PDF结构树 (StructTreeRoot) 和标记信息 (MarkInfo)"},
            {"id": "cleanup_remove_annotations", "title": "清理所有高亮/批注", "desc": "删除文本框、高亮、画笔等所有非链接类型的交互式注释"},
            {"id": "cleanup_remove_metadata", "title": "清空文档元数据", "desc": "移除所有标题、作者、创建时间等PieceInfo和Metadata"},
            {"id": "cleanup_remove_all_links_bookmarks", "title": "移除全部链接和书签", "desc": "仅删除页面链接与书签，不删除普通批注"},
        ],
    },
    {
        "icon": "📦",
        "title": "文件级优化与输出",
        "options": [
            {"id": "convert_pdf_version", "title": "PDF版本转换", "desc": "将PDF版本修改为1.7版本"},
            {"id": "remove_pdf_restrictions", "title": "PDF解除权限限制", "desc": "尝试移除禁止复制、打印、编辑等权限限制，不处理需要打开密码的加密文档"},
            {"id": "fast_web_view", "title": "启用线性化 (快速网页浏览)", "desc": "优化文档结构以支持Web环境下的流式加载和边下边看"},
            {"id": "compress_standard", "title": "标准文件压缩", "desc": "使用更激进的垃圾回收和对象清理，安全减小文件体积（无质量损失）"},
            {"id": "compress_aggressive", "title": "深度文件压缩", "desc": "最大化压缩：garbage=4 + clean模式，适用于接近大小限制的文件"},
            {"id": "compress_images", "title": "压缩内嵌图像 ⚠️", "desc": "将图像降采样至指定DPI，显著减小扫描类PDF体积，但可能影响图像清晰度"},
            {"id": "filename_ectd_format", "title": "eCTD文件名合规格式化", "desc": "自动将输出文件名转为小写、去除空格并替换非法字符"},
        ],
    },
    {
        "icon": "🔍",
        "title": "文档检测",
        "options": [],
    },
]

# 快速预设：key → {title, options}
PRESETS = {
    "china": {
        "title": "中国 eCTD",
        "options": {
            "convert_pdf_version",
            "fast_web_view",
            "initial_view_bookmarks_and_page",
            "page_layout_default",
            "open_page_first",
            "bookmark_inherit_zoom",
            "cleanup_remove_external_uri",
            "cleanup_remove_annotations",
            "cleanup_remove_metadata",
            "cleanup_remove_attachments",
            "cleanup_remove_dynamic_content",
            "link_inherit_zoom",
            "link_open_new_window",
            "bookmark_open_new_window",
            "collapse_all_bookmarks",
            "filename_ectd_format",
        },
    },
    "us": {
        "title": "美国 eCTD",
        "options": {
            "convert_pdf_version",
            "fast_web_view",
            "initial_view_bookmarks_and_page",
            "page_layout_default",
            "open_page_first",
            "bookmark_inherit_zoom",
            "cleanup_remove_annotations",
            "cleanup_remove_metadata",
            "cleanup_remove_attachments",
            "cleanup_remove_dynamic_content",
            "link_inherit_zoom",
            "link_open_new_window",
            "bookmark_open_new_window",
        },
    },
}

# id → 标题 的扁平映射（预检日志、建议列表等场景使用）
OPTION_TITLES = {
    opt["id"]: opt["title"]
    for module in MODULES
    for opt in module["options"]
}


def option_title(option_id):
    """返回规则的显示标题；未知 id 原样返回（与历史行为一致）。"""
    return OPTION_TITLES.get(option_id, option_id)
