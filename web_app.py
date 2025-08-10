import io
import os
import sys
from datetime import datetime
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st


# ====================
# 运行环境自检（避免直接 python 执行导致的上下文缺失）
# ====================
def _ensure_streamlit_context():
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx  # type: ignore
        ctx = get_script_run_ctx()
    except Exception:  # noqa: BLE001
        ctx = None

    if ctx is None:
        # 非 Streamlit 运行，尝试自启动并退出当前进程
        if __name__ == "__main__":
            script_path = os.path.abspath(__file__)
            print("检测到直接运行，正在使用 `streamlit run` 重新启动...")
            # 使用当前 Python 解释器运行，更加稳健（兼容 Windows）
            os.execv(sys.executable, [sys.executable, "-m", "streamlit", "run", script_path])
        else:
            print("请使用命令运行：streamlit run web_app.py")
            sys.exit(0)


# 确保在任何 Streamlit API 调用前完成环境自检
_ensure_streamlit_context()


# ====================
# 常量与配置
# ====================
SUPPORTED_ENCODINGS = ["utf-8", "gbk", "gb2312", "utf-8-sig", "latin1"]
SUPPORTED_FORMATS = {".xlsx": "Excel", ".csv": "CSV"}
OUTPUT_PREFIX = "合并结果_"


# ====================
# 工具函数：读取 CSV（多编码重试）
# ====================
def _read_csv_with_encoding(content: bytes, **kwargs) -> pd.DataFrame:
    last_error: Exception | None = None
    for encoding in SUPPORTED_ENCODINGS:
        try:
            return pd.read_csv(io.BytesIO(content), encoding=encoding, **kwargs)
        except Exception as e:  # noqa: PERF203
            last_error = e
            continue
    raise RuntimeError(f"无法使用常见编码读取CSV。最后错误: {last_error}")


# ====================
# 数据标准化与组装
# ====================
def normalize_columns(dataframes: List[pd.DataFrame], base_columns: List[str]) -> List[pd.DataFrame]:
    normalized: List[pd.DataFrame] = []
    max_columns = len(base_columns)
    for df in dataframes:
        if df is None or df.empty:
            continue
        df = df.dropna(how="all")
        if df.empty:
            continue
        current_cols = len(df.columns)
        if current_cols > max_columns:
            df = df.iloc[:, :max_columns]
        elif current_cols < max_columns:
            for i in range(current_cols, max_columns):
                df[i] = None
        df.columns = base_columns
        normalized.append(df)
    return normalized


def combine_data_with_header(header_df: pd.DataFrame, data_df: pd.DataFrame, base_columns: List[str]) -> pd.DataFrame:
    if header_df is not None and not header_df.empty:
        header_df = header_df.reindex(columns=range(len(base_columns)))
        header_df.columns = base_columns
    return pd.concat([header_df, data_df], ignore_index=True) if header_df is not None else data_df


# ====================
# 读取上传文件为数据帧
# ====================
def read_header_from_upload(file_name: str, content: bytes, skip_rows: int) -> pd.DataFrame:
    if skip_rows == 0:
        return pd.DataFrame()
    if file_name.lower().endswith(".csv"):
        return _read_csv_with_encoding(content, nrows=skip_rows, header=None)
    if file_name.lower().endswith(".xlsx"):
        return pd.read_excel(io.BytesIO(content), nrows=skip_rows, header=None)
    raise ValueError("不支持的文件类型")


def read_data_from_upload(file_name: str, content: bytes, skip_rows: int) -> pd.DataFrame:
    if file_name.lower().endswith(".csv"):
        return _read_csv_with_encoding(content, skiprows=skip_rows, header=None)
    if file_name.lower().endswith(".xlsx"):
        return pd.read_excel(io.BytesIO(content), skiprows=skip_rows, header=None)
    raise ValueError("不支持的文件类型")


def get_base_columns_from_first_file(file_name: str, content: bytes, skip_rows: int) -> List[str]:
    """通过读取数据区首行来获取列数，避免把数据当列名。"""
    if file_name.lower().endswith(".csv"):
        probe = _read_csv_with_encoding(content, skiprows=skip_rows, header=None, nrows=1)
    elif file_name.lower().endswith(".xlsx"):
        probe = pd.read_excel(io.BytesIO(content), skiprows=skip_rows, header=None, nrows=1)
    else:
        raise ValueError("不支持的文件类型")
    num_cols = probe.shape[1]
    return list(range(num_cols))


def write_csv_to_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.StringIO()
    df.to_csv(buffer, index=False, header=False, encoding="utf-8-sig")
    return buffer.getvalue().encode("utf-8-sig")


def write_excel_to_bytes(df: pd.DataFrame, skip_rows: int) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False, header=False)
        if skip_rows > 0:
            try:
                from openpyxl.styles import Font

                worksheet = writer.sheets["Sheet1"]
                bold_font = Font(bold=True)
                for col_idx in range(1, len(df.columns) + 1):
                    cell = worksheet.cell(row=skip_rows, column=col_idx)
                    cell.font = bold_font
            except Exception:  # noqa: BLE001
                # 如果 openpyxl 样式失败，不影响功能
                pass
    buffer.seek(0)
    return buffer.getvalue()


def generate_output_filename(file_count: int, file_format: str, skip_rows: int) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    format_name = SUPPORTED_FORMATS.get(f".{file_format.lower()}", file_format)
    return f"{OUTPUT_PREFIX}{file_count}个{format_name}文件_跳过{skip_rows}行_{timestamp}.{file_format.lower()}"


# ====================
# 合并主逻辑（针对上传文件）
# ====================
def merge_uploaded_files(grouped_files: Dict[str, List[Tuple[str, bytes]]], skip_rows: int):
    results: Dict[str, Dict[str, object]] = {}

    for ext, files in grouped_files.items():
        if not files:
            continue

        files = sorted(files, key=lambda x: x[0])
        first_name, first_content = files[0]

        # 1) 基准列
        base_columns = get_base_columns_from_first_file(first_name, first_content, skip_rows)

        # 2) 标题预览
        header_df = read_header_from_upload(first_name, first_content, skip_rows)
        last_header_row_preview = None
        if skip_rows > 0 and header_df is not None and not header_df.empty:
            last_header_row_preview = header_df.tail(1).values.tolist()[0]

        # 3) 数据读取
        data_frames: List[pd.DataFrame] = []
        for name, content in files:
            df = read_data_from_upload(name, content, skip_rows)
            data_frames.append(df)

        # 4) 标准化
        normalized = normalize_columns(data_frames, base_columns)
        if not normalized:
            results[ext] = {
                "status": "empty",
                "message": f"未找到有效的 {SUPPORTED_FORMATS[ext]} 数据进行合并",
            }
            continue

        # 5) 合并
        combined = pd.concat(normalized, ignore_index=True)
        final_df = combine_data_with_header(header_df, combined, base_columns)

        # 6) 输出
        file_format = ext.lstrip(".")
        if ext == ".csv":
            data_bytes = write_csv_to_bytes(final_df)
        else:  # .xlsx
            data_bytes = write_excel_to_bytes(final_df, skip_rows)
        output_name = generate_output_filename(len(files), file_format, skip_rows)

        results[ext] = {
            "status": "ok",
            "file_name": output_name,
            "bytes": data_bytes,
            "preview": final_df.head(10),
            "rows": len(combined),
            "files": [n for n, _ in files],
            "last_header_row": last_header_row_preview,
            "base_columns_count": len(base_columns),
        }

    return results


# ====================
# Streamlit UI
# ====================
def _inject_minimal_compact_style() -> None:
    """注入轻量紧凑风格的全局 CSS，使默认展示更加简约与节省空间。"""
    st.markdown(
        """
        <style>
        :root {
            /* 色板：有限且和谐 */
            --color-primary: #4f46e5;
            --color-primary-600: #4f46e5;
            --color-accent: #16a34a;
            --color-text: #111827;
            --color-muted: #6b7280;
            --color-border: #e5e7eb;
            --color-surface: #ffffff;
            --color-surface-tint: #f7faff;
            --shadow-sm: 0 2px 8px rgba(17,24,39,0.06);
            --radius-md: 10px;
            --space-1: 0.25rem; /* 4px */
            --space-2: 0.5rem;  /* 8px */
            --space-3: 0.75rem; /* 12px */
            --space-4: 1rem;   /* 16px */
        }

        /* 全局：现代无衬线字体、抗锯齿与紧凑间距 */
        html, body, [data-testid="stAppViewContainer"] {
            font-size: 14px;
            font-family: Inter, "SF Pro Text", "SF Pro Display", "Segoe UI", Roboto, "Helvetica Neue", Arial, "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif;
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
            color: var(--color-text);
        }

        /* 主容器：更紧凑的留白与响应式分布 */
        .main .block-container {
            padding: var(--space-3) var(--space-4);
        }
        /* 列容器（Columns）使用更小的横向间距 */
        [data-testid="stHorizontalBlock"] { gap: var(--space-3) !important; align-items: flex-start; }

        /* 侧边栏上下留白更小 */
        section[data-testid="stSidebar"] > div { padding-top: var(--space-3); padding-bottom: var(--space-3); }

        /* 控件标签与控件之间的间距更紧凑 */
        label[data-testid="stWidgetLabel"] { margin-bottom: var(--space-1); }

        /* 按钮更紧凑的内边距与更圆润的圆角 */
        button[kind="primary"], button[kind="secondary"] {
            padding: 0.4rem 0.8rem;
            border-radius: 8px;
            transition: background-color .15s ease, box-shadow .15s ease, transform .12s ease;
        }

        /* DataFrame 表格更紧凑的行高 */
        [data-testid="stDataFrame"] div[role="row"] { min-height: 24px; }

        /* 折叠面板内容内边距更小 */
        div[role="region"][aria-label="stExpander"] > div {
            padding-top: 0.25rem;
            padding-bottom: 0.25rem;
        }

        /* 标题底部间距更小，整体更简约 */
        h1, h2, h3 {
            margin-bottom: 0.5rem;
            line-height: 1.3;
            letter-spacing: normal;
            overflow: visible;
        }

        /* 卡片式容器优化（对带边框的容器与主要块）*/
        div[data-testid="stContainer"] {
            border-radius: var(--radius-md);
            background: var(--color-surface);
            border: 1px solid var(--color-border);
            box-shadow: var(--shadow-sm);
            transition: box-shadow .15s ease, transform .12s ease;
        }
        div[data-testid="stContainer"]:hover { box-shadow: 0 4px 14px rgba(17,24,39,0.10); }

        /* 选项卡与内容间距更紧凑 */
        div[role="tablist"] { margin-bottom: 0.5rem; }
        [role="tab"] { padding: 0.25rem 0.75rem; color: var(--color-muted); transition: color .15s ease, border-color .15s ease; }
        [role="tab"][aria-selected="true"] {
            border-bottom: 2px solid var(--color-primary);
            color: var(--color-text);
        }

        /* 文件上传区域：中文提示 + 淡色背景突出 */
        [data-testid="stFileUploaderDropzone"] {
            position: relative;
            border: 1px dashed var(--color-border) !important;
            background: var(--color-surface-tint);
            border-radius: var(--radius-md);
            transition: border-color .15s ease, box-shadow .15s ease;
            padding-right: var(--space-4); /* 按钮隐藏后恢复正常内边距 */
        }
        [data-testid="stFileUploaderDropzone"]:hover { border-color: var(--color-primary); box-shadow: var(--shadow-sm); }
        
        [data-testid="stFileUploader"] button:not([aria-label]):hover { filter: brightness(1.04); transform: translateY(-1px); }
        [data-testid="stFileUploader"] button:not([aria-label])::after {
            content: "选择文件";
            position: absolute;
            inset: 0;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #ffffff;
            font-size: 14px; /* 恢复可读字号 */
        }
        /* 还原文件列表内的图标按钮样式，避免文字覆盖 */
        [data-testid="stFileUploader"] button[aria-label] {
            color: inherit;
            background: initial;
            box-shadow: none;
            transform: none;
            filter: none;
        }
        [data-testid="stFileUploader"] button[aria-label]::after { content: none !important; }

        /* 主要按钮更现代的渐变与动效 */
        button[kind="primary"] {
            background: var(--color-primary);
            border: none;
            box-shadow: var(--shadow-sm);
        }
        button[kind="primary"]:hover { filter: brightness(1.04); transform: translateY(-1px); }
        button[kind="primary"]:active { transform: translateY(0); filter: brightness(0.98); }
        button[kind="primary"]:focus { outline: 2px solid rgba(79,70,229,0.35); outline-offset: 2px; }

        /* 轻量“芯片”组件 */
        .chip { display: inline-flex; align-items: center; padding: 2px 8px; border-radius: 999px; background: #f1f5f9; color: #334155; font-size: 12px; border: 1px solid #e2e8f0; margin-right: 6px; }

        /* 顶部 Hero 标题更具层次与高级感 */
        .app-hero { padding: 0.25rem 0 0.5rem 0; }
        .app-hero h1 { font-size: 28px; font-weight: 800; letter-spacing: -0.02em; color: var(--color-text); margin: 0 0 6px 0; }
        .app-hero .subtitle { color: #6b7280; margin: 0 0 8px 0; }
        .app-hero .ver { font-weight: 700; font-size: 0.9em; }
        .app-hero .hero-top { display: flex; align-items: baseline; justify-content: space-between; gap: 12px; }
        .app-hero .author {
            color: #111827;
            background: #eef2ff;
            border: 1px solid #e0e7ff;
            padding: 2px 10px;
            border-radius: 999px;
            font-weight: 700;
            font-size: 12px;
        }
        /* 响应式：在窄屏下增加留白并让表格高度更低 */
        @media (max-width: 1100px) {
            .main .block-container { padding-left: var(--space-3); padding-right: var(--space-3); }
            [data-testid="stDataFrame"] { height: 320px !important; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

st.set_page_config(page_title="Excel/CSV 批量合并工具 v1.0", layout="wide")

_inject_minimal_compact_style()

# 顶部 Hero，现代化标题与标签
with st.container(border=False):
    st.markdown(
        """
        <div class="app-hero">
            <div class="hero-top">
                <h1>Excel/CSV 批量合并工具 <span class="ver">v1.0 沈浪</span></h1>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with st.sidebar:
    st.header("合并设置")
    skip_rows = st.number_input("跳过行数（标题区行数）", min_value=0, value=1, step=1)
    st.divider()
    st.subheader("上传文件")
    uploaded_files = st.file_uploader(
        "选择要合并的文件（可多选）",
        type=["csv", "xlsx"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )
    if uploaded_files:
        st.caption(f"已选择 {len(uploaded_files)} 个文件")
    else:
        st.caption("尚未选择文件")
    st.divider()
    st.subheader("执行合并")
    run = st.button("开始合并", type="primary", use_container_width=True)

left_col, right_col = st.columns([0.01, 0.99], gap="small")

with right_col:
    if run:
        if not uploaded_files:
            st.warning("请先选择文件")
            st.stop()

        # 将上传文件按扩展名分组
        grouped: Dict[str, List[Tuple[str, bytes]]] = {".csv": [], ".xlsx": []}
        for uf in uploaded_files:
            name: str = uf.name
            content: bytes = uf.read()
            uf.seek(0)
            ext = os.path.splitext(name)[1].lower()
            if ext in grouped:
                grouped[ext].append((name, content))

        if not any(grouped.values()):
            st.error("未识别到可合并的 CSV/Excel 文件")
            st.stop()

        with st.spinner("正在合并，请稍候..."):
            results = merge_uploaded_files(grouped, int(skip_rows))

        # 展示与下载：使用选项卡 + 双栏卡片式布局（右侧）
        # 仅展示存在结果的类型，并按 CSV、Excel 顺序
        display_order = [ext for ext in [".csv", ".xlsx"] if ext in results]
        if display_order:
            tabs = st.tabs([SUPPORTED_FORMATS[ext] for ext in display_order])
            for idx, ext in enumerate(display_order):
                with tabs[idx]:
                    result = results.get(ext)
                    if not result:
                        st.info("无可用数据")
                        continue

                    status = result.get("status")
                    if status != "ok":
                        st.info(result.get("message", "无可用数据"))
                        continue

                    files = result["files"]

                    left, right = st.columns([0.36, 0.64])

                    with left:
                        with st.container(border=True):
                            st.subheader("信息与下载")
                            m1, m2, m3 = st.columns(3)
                            m1.metric("文件数", len(files))
                            m2.metric("合并行数", result["rows"])
                            m3.metric("列数", result["base_columns_count"])

                            if result.get("last_header_row") is not None:
                                with st.expander("标题区最后一行预览", expanded=False):
                                    st.write(result["last_header_row"])  # 一行数组

                            with st.expander("已参与合并的文件清单", expanded=False):
                                st.write(files)

                            st.download_button(
                                label="下载合并结果",
                                data=result["bytes"],
                                file_name=result["file_name"],
                                use_container_width=True,
                                mime=(
                                    "text/csv"
                                    if ext == ".csv"
                                    else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                ),
                            )

                    with right:
                        with st.container(border=True):
                            st.subheader("数据预览（前10行）")
                            st.dataframe(result["preview"], use_container_width=True, height=260)

