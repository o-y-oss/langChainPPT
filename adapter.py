"""Only this module depends on ppt-master's command-line and SVG interfaces."""
import queue
import threading
import time
import json
import os
import re
import subprocess
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation

try:
    from langsmith import traceable
except ImportError:
    def traceable(*args, **kwargs):
        def decorator(function):
            return function
        return decorator

TOKEN = re.compile(r"\{\{([^{}\r\n]+)\}\}")


@traceable(run_type="tool")
def run_tool(scripts, logs, name, *args, required=True):
    script = scripts / name

    if not script.is_file():
        raise FileNotFoundError(
            f"未找到 {script}。"
            "请核对 PPT_MASTER_DIR 和仓库版本。"
        )

    command = [
        sys.executable,
        "-X",
        "utf8",
        str(script),
        *map(str, args),
    ]

    timeout_seconds = int(
        os.getenv(
            "PPT_AGENT_TOOL_TIMEOUT",
            "900",
        )
    )

    print(
        f"\n[工具开始] {name}",
        flush=True,
    )
    print(
        "[工具参数]",
        " ".join(map(str, args)),
        flush=True,
    )

    started_at = time.monotonic()
    output_queue = queue.Queue()
    output_finished = threading.Event()

    process = subprocess.Popen(
        command,
        cwd=scripts,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
        env={
            **os.environ,
            "PYTHONUTF8": "1",
            "PYTHONIOENCODING": "utf-8",
        },
    )

    def read_output():
        try:
            if process.stdout is not None:
                for line in process.stdout:
                    output_queue.put(line)
        finally:
            output_finished.set()

    reader = threading.Thread(
        target=read_output,
        daemon=True,
    )
    reader.start()

    collected_output = []
    last_heartbeat = started_at

    log_path = logs / f"{Path(name).stem}.log"

    with log_path.open(
        "a",
        encoding="utf-8",
    ) as log:
        while True:
            try:
                line = output_queue.get(
                    timeout=1
                )
            except queue.Empty:
                now = time.monotonic()
                elapsed = now - started_at

                if (
                    now - last_heartbeat
                    >= 10
                    and process.poll() is None
                ):
                    print(
                        f"[工具运行中] {name} "
                        f"已运行 {elapsed:.0f} 秒，"
                        "暂时没有新的输出",
                        flush=True,
                    )
                    last_heartbeat = now

                if elapsed > timeout_seconds:
                    process.kill()
                    process.wait()

                    raise TimeoutError(
                        f"{name} 运行超过 "
                        f"{timeout_seconds} 秒，"
                        f"已终止。\n"
                        f"详细日志：{log_path}"
                    )

                if (
                    process.poll() is not None
                    and output_finished.is_set()
                    and output_queue.empty()
                ):
                    break

                continue

            collected_output.append(line)
            log.write(line)
            log.flush()

            print(
                f"[{name}] {line.rstrip()}",
                flush=True,
            )

        return_code = process.wait()

    elapsed = time.monotonic() - started_at
    output = "".join(collected_output)

    print(
        f"[工具结束] {name}，"
        f"退出码={return_code}，"
        f"耗时={elapsed:.1f} 秒",
        flush=True,
    )

    if required and return_code != 0:
        raise RuntimeError(
            f"{name} 执行失败，退出码为 {return_code}。\n"
            f"详细日志：{log_path}\n"
            f"最后输出：\n{output[-2000:]}"
        )

    return subprocess.CompletedProcess(
        command,
        return_code,
        stdout=output,
        stderr="",
    )


def svg_slots(folder):
    summary = json.loads((folder / "authoring_summary.json").read_text("utf8"))
    slots = {}
    for doc in summary["documents"]:
        page = doc["file"]
        path = folder / page
        tree = etree.parse(str(path))
        for node in tree.xpath("//*[local-name()='text']"):
            raw = "".join(node.itertext())
            if "{{" not in raw and "}}" not in raw:
                continue
            token = TOKEN.fullmatch(raw)
            leaves = [e for e in node.iter() if e.text]
            refs = node.xpath("ancestor::*[@data-pptx-source-ref][1]/@data-pptx-source-ref")
            if not token or len(leaves) != 1 or leaves[0].text != raw or not refs or not refs[0].startswith("slide:"):
                raise ValueError(f"{page} 中的标记须为页面本地的独立单行、单一样式文本：{raw}")
            key = token[1]
            if key in slots:
                raise ValueError(f"占位符重名：{key}。请在整套模板中使用唯一名称。")
            slots[key] = {"page": page, "label": key, "context": [
                "".join(x.itertext())[:100] for x in tree.xpath("//*[local-name()='text']")
                if "{{" not in "".join(x.itertext())
            ][:15]}
    if not slots:
        raise ValueError("未找到 {{名称}} 占位符，请按说明设置模板。")
    return slots


def apply_texts(folder, originals, slots, edits):
    # Rebuild every changed page from its initial bytes on each repair.
    for page in sorted({s["page"] for s in slots.values()}):
        tree = etree.fromstring(originals[page])
        by_token = {"{{" + e.slot + "}}": e.text for e in edits if slots[e.slot]["page"] == page}
        for node in tree.xpath("//*[local-name()='text']/descendant-or-self::*"):
            if node.text in by_token:
                node.text = by_token[node.text]
        (folder / page).write_bytes(etree.tostring(tree, encoding="utf-8"))


def verify_text(source, output, slots, edits):
    before, after = Presentation(source), Presentation(output)
    if len(before.slides) != len(after.slides):
        raise ValueError("导出前后页数不一致")
    def texts(shapes):
        result = []
        for shape in shapes:
            if shape.has_text_frame:
                result.append(shape.text)
            if hasattr(shape, "shapes"):
                result.extend(texts(shape.shapes))
        return result
    for index, slide in enumerate(after.slides, 1):
        planned = [e.text for e in edits if int(Path(slots[e.slot]["page"]).stem.split("_")[-1]) == index]
        actual = texts(slide.shapes)
        if any(value not in actual for value in planned) or any(TOKEN.search(t) for t in actual):
            raise ValueError(f"第 {index} 页回读失败，检查导出结果及模板标记。")

def run_with_heartbeat(
    operation,
    label,
    interval=10,
):
    """
    在等待模型请求时定期输出心跳。
    """
    result = {}
    error = {}
    finished = threading.Event()

    started_at = time.monotonic()

    def execute():
        try:
            result["value"] = operation()
        except Exception as exc:
            error["value"] = exc
        finally:
            finished.set()

    worker = threading.Thread(
        target=execute,
        daemon=True,
    )
    worker.start()

    while not finished.wait(interval):
        elapsed = time.monotonic() - started_at
        print(
            f"[模型请求中] {label} "
            f"已等待 {elapsed:.0f} 秒",
            flush=True,
        )

    if "value" in error:
        raise error["value"]

    elapsed = time.monotonic() - started_at

    print(
        f"[模型请求完成] {label}，"
        f"耗时 {elapsed:.1f} 秒",
        flush=True,
    )

    return result["value"]
