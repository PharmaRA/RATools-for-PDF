"""处理/预检/检测状态常量。

这些中文字面值同时出现在 worker 信号、controller 状态路由、树节点状态列
和日志导出解析中，是事实上的跨层协议。集中定义避免各处手打字符串漂移。
字面值保持与历史版本一致，保证旧日志导出与用户习惯不受影响。
"""

# —— 批处理状态（ProcessWorker → controller → 日志导出）——
PROCESSING = "正在处理..."
SUCCESS = "处理完成"
FAILURE = "处理失败"
SKIPPED = "已跳过"
STOPPED = "已停止"
NO_MATCH_SKIPPED = "未匹配跳过"

# —— 预检状态（PreCheckWorker）——
PRECHECKING = "预检中..."
PRECHECK_SUGGESTED = "建议处理"
PRECHECK_NO_ACTION = "无需处理"
PRECHECK_FAILED = "预检失败"
PRECHECK_REVIEW = "需要复核"

# —— 只读检测状态（DetectionWorker）——
DETECTING = "正在检测..."
DETECTION_FOUND = "发现问题"
DETECTION_NONE = "未发现"
DETECTION_FAILED = "检测失败"

# —— IO 动作状态（IOActionWorker）——
IO_RUNNING = "正在执行..."
IO_SUCCESS = "操作成功"
IO_FAILURE = "操作失败"

# 一个文件的终态集合：进入这些状态后该文件本轮不再变化
TERMINAL_PROCESS_STATUSES = frozenset({SUCCESS, FAILURE, SKIPPED, STOPPED})

# 各语义分组（供状态→颜色映射与筛选使用）
POSITIVE_STATUSES = frozenset({SUCCESS, IO_SUCCESS, PRECHECK_NO_ACTION, DETECTION_NONE})
NEGATIVE_STATUSES = frozenset({FAILURE, IO_FAILURE, PRECHECK_FAILED, DETECTION_FAILED})
WARNING_STATUSES = frozenset({
    PRECHECK_SUGGESTED,
    PRECHECK_REVIEW,
    DETECTION_FOUND,
    STOPPED,
    SKIPPED,
    NO_MATCH_SKIPPED,
})


def status_semantic(status_text):
    """返回状态的语义分类：'positive' / 'negative' / 'warning' / 'active'。

    'active' 表示进行中或未知状态，UI 端映射为强调色（进行中蓝）。
    """
    if status_text in POSITIVE_STATUSES:
        return "positive"
    if status_text in NEGATIVE_STATUSES:
        return "negative"
    if status_text in WARNING_STATUSES:
        return "warning"
    return "active"
