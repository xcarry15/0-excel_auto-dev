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
            "preview": final_df.head(100),
            "rows": len(combined),
            "files": [n for n, _ in files],
            "last_header_row": last_header_row_preview,
            "base_columns_count": len(base_columns),
        }

    return results


# ====================
# Streamlit UI
# ====================
st.set_page_config(page_title="Excel/CSV 批量合并工具 v1.0", layout="wide")
st.title("Excel/CSV 批量合并工具 v1.0")
st.caption("支持 CSV 与 Excel，同时上传、预览与一键下载合并结果。")

with st.sidebar:
    st.header("合并设置")
    skip_rows = st.number_input("跳过行数（标题区行数）", min_value=0, value=1, step=1)
    st.markdown("- 建议将标题区域行数设为实际表头所在行数\n- Excel 输出将对标题区域最后一行加粗")

uploaded_files = st.file_uploader(
    "选择要合并的文件（可多选）",
    type=["csv", "xlsx"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.success(f"已选择 {len(uploaded_files)} 个文件")

run = st.button("开始合并", type="primary")

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

    # 展示与下载
    for ext, result in results.items():
        if not result:
            continue
        with st.container(border=True):
            st.subheader(f"{SUPPORTED_FORMATS[ext]} 合并结果")
            status = result.get("status")
            if status != "ok":
                st.info(result.get("message", "无可用数据"))
                continue

            files = result["files"]
            st.write(f"文件数: {len(files)} | 合并后数据行数: {result['rows']} | 基准列数: {result['base_columns_count']}")
            if result.get("last_header_row") is not None:
                with st.expander("标题区最后一行预览", expanded=False):
                    st.write(result["last_header_row"])  # 展示为一行数组

            with st.expander("已参与合并的文件清单", expanded=False):
                st.write(files)

            st.markdown("**合并后数据预览（前100行）**")
            st.dataframe(result["preview"], use_container_width=True, height=320)

            st.download_button(
                label="下载合并结果",
                data=result["bytes"],
                file_name=result["file_name"],
                mime=(
                    "text/csv" if ext == ".csv" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )

