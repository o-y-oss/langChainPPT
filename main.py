"""Word -> approved extractive slide plan -> native round-trip PowerPoint."""
import base64
import re
from PIL import Image, ImageOps
import time

from langchain_core.messages import HumanMessage
import shutil
import subprocess

from pptx import Presentation
import argparse
import json
import os
import tempfile
from functools import partial
from pathlib import Path
from typing import TypedDict

from dotenv import load_dotenv
from langchain_openai import ChatOpenAI
from langgraph.graph import END, START, StateGraph
from pydantic import BaseModel, Field

from adapter import (
    TOKEN,
    apply_texts,
    run_tool,
    run_with_heartbeat,
    svg_slots,
    verify_text,
)

ROOT = Path(__file__).resolve().parent


class Edit(BaseModel):
    slot: str
    source: int = Field(ge=1)
    text: str = Field(min_length=1)


class Plan(BaseModel):
    edits: list[Edit]


class State(TypedDict, total=False):
    source: dict[int, str]
    slots: dict
    plan: Plan
    feedback: str
    tries: int
    failed_pages: list[str]
    valid: bool
    output: str

class OutlineItem(BaseModel):
    order: int = Field(ge=1)
    title: str = Field(min_length=1)
    subpoints: list[str] = Field(default_factory=list)


class OutlineResult(BaseModel):
    items: list[OutlineItem] = Field(min_length=1)

def write_json(path, data):
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf8")


def validate(plan, state):
    keys = [e.slot for e in plan.edits]
    if sorted(keys) != sorted(state["slots"]):
        raise ValueError("计划必须恰好覆盖模板中的每个标记一次")
    for e in plan.edits:
        if (not e.text.strip() or e.text != e.text.strip() or "\n" in e.text
                or "\r" in e.text or TOKEN.search(e.text)
                or e.text not in state["source"].get(e.source, "")):
            raise ValueError(f"{e.slot} 需要填写来源行中的一段连续原文")

def read_reference_ppt(path):
    """返回 (页码, 文字)。图片中嵌入的文字不在本函数的提取范围内。"""
    def read_shapes(shapes):
        for shape in shapes:
            if shape.has_table:
                for row in shape.table.rows:
                    cells = [cell.text.strip() for cell in row.cells]
                    if any(cells):
                        yield " | ".join(cells)
            elif shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    text = paragraph.text.strip()
                    if text:
                        yield text
            elif hasattr(shape, "shapes"):
                yield from read_shapes(shape.shapes)

    result = []
    for page_no, slide in enumerate(Presentation(path).slides, start=1):
        for text in read_shapes(slide.shapes):
            result.append((page_no, text))

    if not result:
        raise ValueError(
            "旧版 PPT 未提取到文字。如果整份 PPT 由截图组成，需要先做 OCR。"
        )
    return result

def prepare_outline_tiles(image_path, run_dir):
    """
    将大纲长图转换为多张符合 DeepSeek Vision 限制的 JPEG 图片。
    """

    from pathlib import Path

    image_path = Path(image_path).resolve()
    run_dir = Path(run_dir).resolve()

    tile_dir = run_dir / "outline_tiles"
    tile_dir.mkdir(parents=True, exist_ok=True)

    try:
        with Image.open(image_path) as source:
            source.load()

            image = ImageOps.exif_transpose(source)

            # 统一转换成 RGB，避免 RGBA、CMYK 等格式导致请求失败
            if "A" in image.getbands():
                rgba = image.convert("RGBA")
                background = Image.new(
                    "RGB",
                    rgba.size,
                    (255, 255, 255),
                )
                background.paste(
                    rgba,
                    mask=rgba.getchannel("A"),
                )
                image = background
            else:
                image = image.convert("RGB")

    except Exception as exc:
        raise RuntimeError(
            f"无法读取大纲图片：{image_path}"
        ) from exc

    # DeepSeek 单边最大尺寸为 8192 像素
    max_side = 8192

    # 如果图片宽度超过限制，先按比例缩小
    if image.width > max_side:
        scale = max_side / image.width
        new_size = (
            max_side,
            max(1, round(image.height * scale)),
        )
        image = image.resize(
            new_size,
            Image.Resampling.LANCZOS,
        )

    # 每张切片的高度低于 8192，留出安全余量
    tile_height = 7600
    overlap = 100
    step = tile_height - overlap

    tiles = []
    start_y = 0
    tile_index = 1

    while start_y < image.height:
        end_y = min(
            start_y + tile_height,
            image.height,
        )

        tile = image.crop(
            (
                0,
                start_y,
                image.width,
                end_y,
            )
        )

        tile_path = (
            tile_dir
            / f"outline_tile_{tile_index:02d}.jpg"
        )

        quality = 82

        while True:
            tile.save(
                tile_path,
                format="JPEG",
                quality=quality,
                optimize=True,
                progressive=True,
            )

            # 每个切片尽量控制在 3 MB 以内
            if tile_path.stat().st_size <= 3 * 1024 * 1024:
                break

            if quality > 50:
                quality -= 10
                continue

            # 降低质量后仍然过大，再缩小尺寸
            tile = tile.resize(
                (
                    max(1, round(tile.width * 0.8)),
                    max(1, round(tile.height * 0.8)),
                ),
                Image.Resampling.LANCZOS,
            )
            quality = 70

        tiles.append(tile_path)

        if end_y >= image.height:
            break

        start_y += step
        tile_index += 1

    if not tiles:
        raise RuntimeError(
            "没有生成有效的大纲图片切片"
        )

    total_size = sum(
        path.stat().st_size
        for path in tiles
    )

    # Base64 编码后体积还会增加，因此预留安全空间
    if total_size > 30 * 1024 * 1024:
        raise RuntimeError(
            "大纲图片切片后的总大小超过 30 MB。"
            "请先将原始长图手动拆成两张或多张图片。"
        )

    print(
        f"大纲图片已切分为 {len(tiles)} 张 JPEG，"
        f"总大小约为 {total_size / 1024 / 1024:.2f} MB"
    )

    return tiles

def read_outline_image(image_path, run_dir):
    """使用 DeepSeek 多模态模型读取大纲图片，确认由工作流节点完成。"""

    import base64
    import json
    import os
    import re
    from pathlib import Path

    from PIL import Image, ImageOps
    from langchain_core.messages import HumanMessage
    from langchain_openai import ChatOpenAI

    image_path = Path(image_path).resolve()
    run_dir = Path(run_dir).resolve()

    if not image_path.is_file():
        raise FileNotFoundError(
            f"大纲图片不存在：{image_path}"
        )

    api_key = os.getenv("DEEPSEEK_API_KEY")
    if not api_key:
        raise RuntimeError(
            "未设置 DEEPSEEK_API_KEY，无法识别大纲图片。"
        )

    base_url = os.getenv(
        "DEEPSEEK_BASE_URL",
        "https://api.deepseek.com",
    )

    vision_model = os.getenv(
        "DEEPSEEK_VISION_MODEL",
        "deepseek-flash",
    )

    vision_tiles = prepare_outline_tiles(
        image_path,
        run_dir,
    )

    image_blocks = []

    for tile_path in vision_tiles:
        encoded_image = base64.b64encode(
            tile_path.read_bytes()
        ).decode("ascii")

        image_blocks.append(
            {
                "type": "image_url",
                "image_url": {
                    "url": (
                        "data:image/jpeg;base64,"
                        + encoded_image
                    )
                },
            }
        )

    model = ChatOpenAI(
        model=vision_model,
        api_key=api_key,
        base_url=base_url,
        timeout=180,
        max_retries=1,
    )

    prompt = """
你负责识别一张中文 PPT 制作大纲图片。

请完成以下任务：

1. 按图片中的真实阅读顺序识别所有章节、页面主题和制作要求。
2. 保留图片中的编号、章节顺序和关键措辞。
3. 不要修改公司名称、产品名称、年份、数字等专有信息。
4. 如果章节下方存在子要点，将其放入 subpoints。
5. 不要补充图片中没有出现的信息。
6. 只返回 JSON，不要返回解释，不要使用 Markdown 代码块。

返回格式：

{
  "items": [
    {
      "order": 1,
      "title": "第一个章节或页面主题",
      "subpoints": ["子要点1", "子要点2"]
    }
  ]
}
"""

    last_error = ""

    for attempt in range(3):
        current_prompt = prompt

        if last_error:
            current_prompt += (
                "\n上一次输出存在以下问题，请修正后重新输出：\n"
                + last_error
            )

        message_content = [
            {
                "type": "text",
                "text": (
                        current_prompt
                        + "\n下面的图片是同一张大纲长图的连续切片，"
                          "图片已经按照从上到下的顺序排列。"
                          "请结合所有图片识别完整大纲。"
                ),
            }
        ]

        message_content.extend(image_blocks)

        message = HumanMessage(
            content=message_content
        )

        response = run_with_heartbeat(
            lambda: model.invoke([message]),
            "大纲多模态识别",
        )

        if isinstance(response.content, list):
            raw_text = "".join(
                block.get("text", "")
                for block in response.content
                if isinstance(block, dict)
            )
        else:
            raw_text = str(response.content)

        raw_text = raw_text.strip()

        # 去掉 ```json ... ``` 包裹
        raw_text = re.sub(
            r"^```(?:json)?\s*",
            "",
            raw_text,
            flags=re.IGNORECASE,
        )
        raw_text = re.sub(
            r"\s*```$",
            "",
            raw_text,
        ).strip()

        # 从模型回答中提取 JSON 对象
        match = re.search(
            r"\{.*\}",
            raw_text,
            flags=re.DOTALL,
        )

        if not match:
            last_error = "模型返回内容中没有找到 JSON 对象"
            continue

        try:
            result = json.loads(match.group(0))

            items = result.get("items")

            if not isinstance(items, list) or not items:
                raise ValueError(
                    "JSON 中的 items 不是非空列表"
                )

            cleaned_items = []

            for index, item in enumerate(items, start=1):
                if not isinstance(item, dict):
                    raise ValueError(
                        f"第 {index} 项不是对象"
                    )

                title = str(
                    item.get("title", "")
                ).strip()

                if not title:
                    raise ValueError(
                        f"第 {index} 项标题为空"
                    )

                subpoints = item.get(
                    "subpoints",
                    [],
                )

                if not isinstance(subpoints, list):
                    subpoints = []

                subpoints = [
                    str(point).strip()
                    for point in subpoints
                    if str(point).strip()
                ]

                cleaned_items.append(
                    {
                        "order": index,
                        "title": title,
                        "subpoints": subpoints,
                    }
                )

            details_path = run_dir / "outline_details.json"
            details_path.write_text(
                json.dumps(
                    {
                        "items": cleaned_items,
                    },
                    ensure_ascii=False,
                    indent=2,
                ),
                encoding="utf-8",
            )

            outline_file = run_dir / "outline_ocr.txt"

            outline_file.write_text(
                "\n".join(
                    f'{item["order"]}. {item["title"]}'
                    + (
                        "｜子要点："
                        + "；".join(item["subpoints"])
                        if item["subpoints"]
                        else ""
                    )
                    for item in cleaned_items
                ),
                encoding="utf-8",
            )

            print(
                f"\n多模态模型识别结果：{outline_file}"
            )
            print(
                outline_file.read_text(
                    encoding="utf-8"
                )
            )

            # 大纲确认由 style_deck.py 的 review_outline 节点负责。
            # 这样命令行模式仍然可以编辑文件确认，LangSmith 模式可以使用 interrupt 恢复。
            final_outline = []

            for line in outline_file.read_text(
                encoding="utf-8-sig"
            ).splitlines():
                line = line.strip()

                if not line:
                    continue

                # 去掉行首编号
                line = re.sub(
                    r"^\s*\d+[\.\、．]\s*",
                    "",
                    line,
                )

                # 后续流程只使用章节标题
                if "｜子要点：" in line:
                    line = line.split(
                        "｜子要点：",
                        1,
                    )[0].strip()

                if line:
                    final_outline.append(line)

            if not final_outline:
                raise ValueError(
                    "outline_ocr.txt 中没有有效的大纲内容"
                )

            outline_json = run_dir / "outline.json"
            outline_json.write_text(
                json.dumps(
                    final_outline,
                    ensure_ascii=False,
                    indent=2,
                ),
                encoding="utf-8",
            )

            return final_outline

        except (ValueError, TypeError, json.JSONDecodeError) as exc:
            last_error = str(exc)

    raise RuntimeError(
        "多模态模型连续三次未能返回有效的大纲 JSON。"
        f"最后一次错误：{last_error}"
    )



def build_langsmith_graph(
    *,
    word: Path,
    template: Path,
    reference_ppt: Path,
    outline_image: Path,
    request: str,
    review_mode: str = "interrupt",
):
    """创建可由 LangGraph Agent Server 暴露的完整风格参考工作流。

    该函数使用与命令行正常模式相同的资料解析、风格分析、内容规划、
    人工审阅、SVG 生成、质量检查和 PPTX 导出逻辑。
    """
    load_dotenv(ROOT / ".env")

    word = Path(word).resolve()
    template = Path(template).resolve()
    reference_ppt = Path(reference_ppt).resolve()
    outline_image = Path(outline_image).resolve()

    for label, path in [
        ("Word 资料", word),
        ("风格参考 PPT", template),
        ("旧版公司介绍 PPT", reference_ppt),
        ("制作大纲长图", outline_image),
    ]:
        if not path.is_file():
            raise FileNotFoundError(f"{label}不存在：{path}")

    if reference_ppt == template:
        raise ValueError("旧版公司介绍 PPT 和风格参考 PPT 不能是同一个文件")

    runs = ROOT / "runs"
    scripts = (ROOT / os.getenv("PPT_MASTER_DIR", "vendor/ppt-master")).resolve() / "skills/ppt-master/scripts"

    def tool_for_run(run):
        return partial(run_tool, scripts, run / "logs")

    def ingest(state):
        # 恢复审阅时从 checkpoint 取回此目录。新任务使用独立目录。
        runs.mkdir(exist_ok=True)
        run = Path(tempfile.mkdtemp(prefix="deck_", dir=runs))
        print("本次目录：", run, flush=True)
        (run / "logs").mkdir()
        work = run / "workspace"
        work.mkdir()
        tool = tool_for_run(run)
        tool("attribution_guard.py")

        md = run / "source.md"
        tool("source_to_md.py", word, "-o", md)

        lines = [
            item.strip()
            for item in md.read_text("utf-8-sig").splitlines()
            if item.strip()
        ]

        if not lines:
            raise ValueError("Word 未提取到有效正文")
        if sum(map(len, lines)) > 25000:
            raise ValueError(
                "Word 转换后的正文超过 25000 字符，请先精简资料。"
            )

        source = {}
        for line_no, text in enumerate(lines, start=1):
            source[len(source) + 1] = {
                "kind": "word",
                "file": word.name,
                "location": f"转换后第 {line_no} 条",
                "text": text,
            }

        old_items = read_reference_ppt(reference_ppt)
        for page_no, text in old_items:
            source[len(source) + 1] = {
                "kind": "reference_ppt",
                "file": reference_ppt.name,
                "location": f"第 {page_no} 页",
                "text": text,
            }

        outline = read_outline_image(outline_image, run)
        write_json(run / "evidence.json", source)
        write_json(run / "outline.json", outline)

        print(
            f"已读取 Word {len(lines)} 条、旧 PPT {len(old_items)} 条。",
            flush=True,
        )
        print(
            f"已识别大纲 {len(outline)} 项，等待工作流确认。",
            flush=True,
        )

        return {
            "source": source,
            "outline": outline,
            "run_dir": str(run),
            "work_dir": str(work),
            "request": state.get("request", request),
            "review_mode": state.get("review_mode", review_mode),
        }

    from style_deck import build_style_graph

    return build_style_graph(
        ingest=ingest,
        tool_for_run=tool_for_run,
        template=template,
        request=request,
        default_review_mode=review_mode,
    )


def build_langsmith_graph_from_env():
    """按照 .env 中的本地资料路径创建 LangSmith 图。"""
    load_dotenv(ROOT / ".env")
    return build_langsmith_graph(
        word=Path(
            os.getenv("PPT_WORD_PATH", ROOT / "input/source.docx")
        ),
        template=Path(
            os.getenv("PPT_TEMPLATE_PATH", ROOT / "input/template.pptx")
        ),
        reference_ppt=Path(
            os.getenv("PPT_REFERENCE_PATH", ROOT / "input/reference.pptx")
        ),
        outline_image=Path(
            os.getenv("PPT_OUTLINE_PATH", ROOT / "input/outline.png")
        ),
        request=os.getenv(
            "PPT_REQUEST",
            "根据资料制作中文公司介绍 PPT，沿用风格参考 PPT 的视觉风格。",
        ),
        review_mode="interrupt",
    )

def main():
    load_dotenv(ROOT / ".env")
    parser = argparse.ArgumentParser()
    parser.add_argument("--word", type=Path, default=ROOT / "input/source.docx")
    parser.add_argument("--template", type=Path, default=ROOT / "input/template.pptx")
    parser.add_argument("--demo", action="store_true", help="使用预设示例计划，提出自然语言修改时才调用模型")
    parser.add_argument(
        "--request",
        default="根据 Word 资料制作中文演示稿。自行规划页数、章节和布局，沿用参考 PPT 的视觉风格。",
    )
    parser.add_argument(
        "--reference-ppt",
        type=Path,
        help="旧版公司介绍 PPTX，用于补充 Word 未覆盖的事实",
    )
    parser.add_argument(
        "--outline-image",
        type=Path,
        help="制作大纲长图，例如 PNG 或 JPG",
    )
    args = parser.parse_args()
    word, template = args.word.resolve(), args.template.resolve()
    if not args.demo:
        for label, path in [
            ("旧版公司介绍 PPT", args.reference_ppt),
            ("制作大纲长图", args.outline_image),
        ]:
            if path is None or not path.is_file():
                parser.error(f"请提供有效的{label}路径，当前值为：{path}")

        if args.reference_ppt.resolve() == template:
            parser.error("旧版公司介绍 PPT 和风格参考 PPT 指向了同一个文件，请核对参数。")
    if not word.is_file() or not template.is_file():
        parser.error("请核对 Word 和模板路径")
    if not args.demo:
        graph = build_langsmith_graph(
            word=word, template=template, reference_ppt=args.reference_ppt,
            outline_image=args.outline_image, request=args.request, review_mode="cli",
        )
        result = graph.invoke({"request": args.request}, {"recursion_limit": 1000})
        if result.get("output"):
            print("PPTX 已导出：", result["output"])
        return

    scripts = (ROOT / os.getenv("PPT_MASTER_DIR", "vendor/ppt-master")).resolve() / "skills/ppt-master/scripts"
    runs = ROOT / "runs"
    runs.mkdir(exist_ok=True)
    run = Path(tempfile.mkdtemp(prefix="deck_", dir=runs))
    print("本次目录：", run, flush=True)
    logs = run / "logs"
    logs.mkdir()
    work = run / "workspace"
    folder = work / "authoring-svg-flat"
    originals = {}
    def tool(name, *params, required=True):
        return run_tool(scripts, logs, name, *params, required=required)

    def ingest(state):
        tool("attribution_guard.py")
        md = run / "source.md"
        tool("source_to_md.py", word, "-o", md)
        lines = [s.strip() for s in md.read_text("utf-8-sig").splitlines() if s.strip()]
        if not lines or sum(map(len, lines)) > 25000:
            raise ValueError("首版支持含可提取正文、总长度不超过 25000 字符的 Word，请精简后重试。")
        source = dict(enumerate(lines, 1))
        write_json(run / "evidence.json", source)
        return {"source": source}

    def inspect_template(state):
        tool("pptx_to_svg.py", template, "-o", work, "--inheritance-mode", "both", "--roundtrip")
        slots = svg_slots(folder)
        for s in slots.values():
            page = s["page"]
            originals[page] = (folder / page).read_bytes()
        write_json(run / "slots.json", slots)
        return {"slots": slots}

    def plan_content(state):
        p = Plan.model_validate({"edits": [dict(slot=f"内容{i}", source=i, text=t)
            for i,t in enumerate(["星河示例公司","公司专注于制造业软件服务。","服务包括设备监控与生产数据分析。"], 1)]})
        validate(p, state)
        return {"plan": p}

    def review(state):
        path = run / "plan.json"
        from style_deck import ask_json, make_text_model, save_json

        plan = state["plan"]
        save_json(path, plan)
        while True:
            print("\n计划文件：", path)
            for e in plan.edits:
                print(f'{state["slots"][e.slot]["page"]} | {e.slot} ← 来源 {e.source} | {e.text}')
            answer = input("输入修改意见；回车或“确认”继续；e 重读文件；q 退出：").strip()
            if answer.casefold() in {"q", "退出", "取消"}:
                raise SystemExit("已停止，计划已保存。")
            if answer.casefold() in {"", "确认", "通过", "同意", "继续", "approve"}:
                return {"plan": plan}
            try:
                if answer.casefold() == "e":
                    updated = Plan.model_validate_json(path.read_text("utf-8-sig"))
                    validate(updated, state)
                else:
                    updated = ask_json(
                        make_text_model(), Plan,
                        "按用户意见修改现有计划，只改涉及的槽位。每个槽位取自对应来源行的连续原文。",
                        {"制作要求": args.request, "当前计划": plan.model_dump(), "修改意见": answer,
                         "槽位": state["slots"], "来源行": state["source"]},
                        lambda candidate: validate(candidate, state), label="示例计划修订",
                    )
            except Exception as exc:
                print(f"修订未完成，继续审阅上一版：{exc}")
                continue
            save_json(path, updated)
            plan = updated

    def fill(state):
        apply_texts(folder, originals, state["slots"], state["plan"].edits)
        tool("svg_authoring_view.py", folder, "--refresh-summary")
        return {}

    def check(state):
        result = tool("svg_quality_checker.py", work, "--roundtrip", "--json", required=False)
        report = work / "validation/svg_quality_report.json"
        if not report.exists():
            raise RuntimeError("检查器未生成报告，请看 logs。")
        print(result.stdout[-2000:])
        data = json.loads(report.read_text("utf8"))
        failed_pages = [row["file"] for row in data["files"] if row.get("errors")]
        if result.returncode and not failed_pages:
            raise RuntimeError(f"检查器执行异常或项目级错误，请看 {report}")
        if result.returncode and (args.demo or state.get("tries", 0) >= 2):
            raise RuntimeError(f"检查未通过，已停止导出。详情：{report}")
        return {"valid": result.returncode == 0, "tries": state.get("tries", 0) + 1,
                "failed_pages": failed_pages,
                "feedback": json.dumps([{"page": row["file"], "errors": row["errors"]}
                    for row in data["files"] if row.get("errors")], ensure_ascii=False)}

    def export(state):
        tool("svg_to_pptx.py", work, "--roundtrip")
        files = list((work / "exports").glob("*.pptx"))
        if len(files) != 1:
            raise RuntimeError("导出数量异常，请查看导出日志")
        out = files[0]
        delivery = json.loads(tool("pptx_delivery_check.py", out).stdout)
        write_json(run / "delivery.json", delivery)
        print("交付检查：", delivery["status"], "；提示数：", len(delivery.get("advisories", [])))
        verify_text(template, out, state["slots"], state["plan"].edits)
        return {"output": str(out)}

    def node_with_progress(name, function):
        def wrapped(state):
            started_at = time.monotonic()

            print(
                f"\n[阶段开始] {name}",
                flush=True,
            )

            try:
                result = function(state)
            except Exception as exc:
                elapsed = time.monotonic() - started_at
                print(
                    f"[阶段失败] {name}，"
                    f"耗时 {elapsed:.1f} 秒，"
                    f"错误：{exc}",
                    flush=True,
                )
                raise
            else:
                elapsed = time.monotonic() - started_at
                print(
                    f"[阶段完成] {name}，"
                    f"耗时 {elapsed:.1f} 秒",
                    flush=True,
                )
                return result

        return wrapped

    g = StateGraph(State)

    nodes = [
        ingest,
        inspect_template,
        plan_content,
        review,
        fill,
        check,
        export,
    ]

    for fn in nodes:
        g.add_node(
            fn.__name__,
            node_with_progress(
                fn.__name__,
                fn,
            ),
        )
    g.add_edge(START, "ingest")
    for current, following in zip(nodes[:4], nodes[1:5]):
        g.add_edge(current.__name__, following.__name__)
    g.add_edge("fill","check")
    g.add_conditional_edges("check", lambda s: "export" if s["valid"] else "plan_content")
    g.add_edge("export", END)
    result = g.compile().invoke({"tries": 0}, {"recursion_limit": 40})
    print("验收通过，PPTX：", result["output"])


if __name__ == "__main__":
    main()
