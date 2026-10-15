#!/usr/bin/env python3
"""Audit structure, audible counts, and repeated cut-point shots in a dialogue-driven storyboard draft."""

from __future__ import annotations

import argparse
import re
import sys
import unicodedata
from collections import Counter
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path


SECTION_RE = re.compile(
    r"^(?P<prefix>## )?视频编号(?P<num>\d+)（(?P<title>.+?)，(?P<duration>\d+(?:\.\d+)?)秒 / 共(?P<count>\d+)镜）\s*$", re.M
)
SHOT_RE = re.compile(
    r"^镜头(?P<num>\d+) / SC1（(?P<start>\d+(?:\.\d+)?)—(?P<end>\d+(?:\.\d+)?)秒）\s*$",
    re.M,
)
FIELDS = [
    "时间标识",
    "摄影机位置",
    "运镜",
    "画面",
    "动作/情绪",
    "位置承接",
    "OS",
    "画外音",
    "音效",
    "台词",
]
SOUND_FIELDS = ["OS", "画外音", "台词"]
CAMERA_META_RE = re.compile(r"上一(?:编号|镜)|前一(?:编号|镜)|摄影机不越过|同(?:一)?侧|保持|原轴线|(?:发言者|说话者|听者|对话者)")
RELATIVE_VISUAL_RE = re.compile(r"上一(?:编号|镜)|前一(?:编号|镜)|同前|如前|仍在原处|动作继续|(?:保持|沿用|延续).{0,8}(?:上一|前一|原)(?:镜|编号|机位|构图|位置|轴线)?")


@dataclass
class Finding:
    level: str
    section: str
    message: str

    def render(self) -> str:
        return f"{self.level} {self.section}: {self.message}"


def audible_count(text: str) -> int:
    """Count spoken letters and numbers; ignore punctuation, symbols, and spacing."""
    return sum(1 for char in text if unicodedata.category(char)[0] in {"L", "N"})


def field_value(block: str, field: str) -> str | None:
    match = re.search(rf"^【{re.escape(field)}】(.*)$", block, re.M)
    return match.group(1).strip() if match else None


def quoted_spoken_text(value: str) -> str:
    parts = re.findall(r"[“\"](.*?)[”\"]", value)
    return "".join(parts)


def visual_value(block: str, field: str) -> str:
    return re.sub(r"\s+", "", field_value(block, field) or "")


def visual_similarity(first: str, second: str) -> float:
    return SequenceMatcher(None, first, second).ratio() if first and second else 0.0


def movement_family(value: str | None) -> str | None:
    if not value or value in {"固定镜头", "固定机位", "静止镜头"}:
        return None
    for family, pattern in (
        ("变焦", r"变焦|放大|缩放"),
        ("弧移", r"弧移|环绕|绕拍"),
        ("跟拍", r"跟拍|跟随"),
        ("横移", r"横移|侧移|平移"),
        ("摇镜", r"摇镜|轻摇|横摇|俯仰|上摇|下摇"),
        ("推近", r"推近|缓推|推进|推镜"),
        ("拉远", r"拉远|缓拉|后拉|拉镜"),
    ):
        if re.search(pattern, value):
            return family
    return "其他运镜"


def audit(text: str) -> list[Finding]:
    findings: list[Finding] = []
    matches = list(SECTION_RE.finditer(text))
    if not matches:
        return [Finding("ERROR", "全文", "没有找到视频编号标题")]

    nums = [int(match.group("num")) for match in matches]
    for previous, current in zip(nums, nums[1:]):
        if current != previous + 1:
            findings.append(Finding("ERROR", "全文", f"编号不连续：{previous} 后为 {current}"))

    previous_boundary: tuple[str, dict[str, str]] | None = None
    total_shots = 0
    moving_shots = 0
    movement_counts: Counter[str] = Counter()
    for index, match in enumerate(matches):
        label = f"视频编号{match.group('num')}"
        duration = float(match.group("duration"))
        declared_shots = int(match.group("count"))
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        section = text[match.start():end]
        if not match.group("prefix"):
            findings.append(Finding("WARN", label, "标题缺少格式要求的## 前缀"))
        if not 0 < duration <= 15:
            findings.append(Finding("ERROR", label, f"标题时长{duration:g}秒不在允许范围内"))
        if duration < 13:
            findings.append(Finding("WARN", label, "短于通常的13—15秒；检查剧情与动作时长"))

        scene_match = re.search(r"^【本节固定场景】(.+)$", section, re.M)
        if not scene_match:
            findings.append(Finding("ERROR", label, "缺少【本节固定场景】"))
        else:
            scene_text = scene_match.group(1).strip()
            scenes = [item.strip() for item in re.split(r"[，,、;；]", scene_text) if item.strip()]
            if len(scenes) > 2:
                findings.append(Finding("ERROR", label, f"固定场景超过两个：{len(scenes)}"))
            if "相连" in scene_text or "连接" in scene_text:
                findings.append(Finding("ERROR", label, "固定场景使用了连接式描述"))

        cast_match = re.search(r"^【本节出场人物】(.+)$", section, re.M)
        if not cast_match:
            findings.append(Finding("ERROR", label, "缺少【本节出场人物】"))
        else:
            cast_text = cast_match.group(1).strip()
            cast = [item.strip() for item in re.split(r"[，,、;；]", cast_text) if item.strip()]
            if len(cast) > 4:
                findings.append(Finding("ERROR", label, f"角色单位超过四个：{len(cast)}"))
            if cast_text.count("群像") > 1:
                findings.append(Finding("ERROR", label, "群像超过一个"))

        declared_match = re.search(r"^【本节有声字数】(\d+)字\s*$", section, re.M)
        if not declared_match:
            findings.append(Finding("ERROR", label, "缺少或错误填写【本节有声字数】"))

        shots = list(SHOT_RE.finditer(section))
        if not shots:
            findings.append(Finding("ERROR", label, "没有找到镜头"))
            previous_boundary = None
            continue
        total_shots += len(shots)
        if len(shots) != declared_shots:
            findings.append(Finding("ERROR", label, f"标题声明{declared_shots}镜，实际{len(shots)}镜"))

        shot_nums = [int(shot.group("num")) for shot in shots]
        if shot_nums != list(range(1, len(shots) + 1)):
            findings.append(Finding("ERROR", label, f"镜头编号错误：{shot_nums}"))

        ranges = [(float(shot.group("start")), float(shot.group("end"))) for shot in shots]
        if abs(ranges[0][0]) > 1e-9:
            findings.append(Finding("ERROR", label, "镜头1没有从0秒开始"))
        if abs(ranges[-1][1] - duration) > 1e-9:
            findings.append(Finding("ERROR", label, f"末镜没有在{duration:g}秒结束"))
        for shot_index, (start, stop) in enumerate(ranges, 1):
            if stop <= start:
                findings.append(Finding("ERROR", label, f"镜头{shot_index}时长非正数"))
        for shot_index, ((_, stop), (next_start, _)) in enumerate(zip(ranges, ranges[1:]), 1):
            if abs(stop - next_start) > 1e-9:
                findings.append(
                    Finding("ERROR", label, f"镜头{shot_index}结束{stop:g}秒与下一镜开始{next_start:g}秒不连续")
                )

        computed = 0
        voiced_shots = 0
        shot_blocks: list[str] = []
        section_movements: list[str | None] = []
        for shot_index, shot in enumerate(shots):
            shot_end = shots[shot_index + 1].start() if shot_index + 1 < len(shots) else len(section)
            block = section[shot.start():shot_end]
            shot_blocks.append(block)
            family = movement_family(field_value(block, "运镜"))
            section_movements.append(family)
            if family:
                moving_shots += 1
                movement_counts[family] += 1
            positions = []
            for field in FIELDS:
                field_match = re.search(rf"^【{re.escape(field)}】", block, re.M)
                if not field_match:
                    findings.append(Finding("ERROR", label, f"镜头{shot_index + 1}缺少【{field}】"))
                else:
                    positions.append((field, field_match.start()))
            if len(positions) == len(FIELDS):
                actual_order = [name for name, _ in sorted(positions, key=lambda item: item[1])]
                if actual_order != FIELDS:
                    findings.append(Finding("ERROR", label, f"镜头{shot_index + 1}栏目顺序错误"))

            for field in ("时间标识", "运镜", "画面", "动作/情绪", "位置承接"):
                value = field_value(block, field) or ""
                if RELATIVE_VISUAL_RE.search(value):
                    findings.append(Finding("WARN", label, f"镜头{shot_index + 1}【{field}】依赖其他镜头；写出本镜可独立理解的具体状态"))

            camera_position = field_value(block, "摄影机位置") or ""
            if CAMERA_META_RE.search(camera_position):
                findings.append(Finding("WARN", label, f"镜头{shot_index + 1}【摄影机位置】含规划或泛指用语；改写为相对具体人物或场景实体的当前机位、景别与主拍对象"))

            active: list[tuple[str, str]] = []
            for field in SOUND_FIELDS:
                value = field_value(block, field)
                if value is not None and value != "无":
                    active.append((field, value))
            if len(active) > 1:
                findings.append(Finding("ERROR", label, f"镜头{shot_index + 1}存在多个声音来源"))
            if active:
                voiced_shots += 1
                spoken = quoted_spoken_text(active[0][1])
                if not spoken:
                    findings.append(Finding("ERROR", label, f"镜头{shot_index + 1}有声正文缺少引号"))
                spoken_count = audible_count(spoken)
                computed += spoken_count
                if spoken_count > 20:
                    findings.append(Finding("ERROR", label, f"镜头{shot_index + 1}单镜有声{spoken_count}字超过20；按原文语义拆成接续镜头"))
            elif ranges[shot_index][1] - ranges[shot_index][0] > 3:
                findings.append(Finding("WARN", label, f"镜头{shot_index + 1}无声超过3秒；缩短无剧情信息的停留或调整分镜，保留关键画面"))

        run_family: str | None = None
        run_start = 0
        for shot_index, family in enumerate(section_movements + [None]):
            if family != run_family:
                if run_family and shot_index - run_start >= 3:
                    findings.append(Finding("WARN", label, f"镜头{run_start + 1}—{shot_index}连续使用{run_family}；逐镜核对画面任务与运动结果"))
                run_family = family
                run_start = shot_index

        for shot_index in range(1, len(shot_blocks)):
            earlier, later = shot_blocks[shot_index - 1], shot_blocks[shot_index]
            camera_a, camera_b = visual_value(earlier, "摄影机位置"), visual_value(later, "摄影机位置")
            picture_a, picture_b = visual_value(earlier, "画面"), visual_value(later, "画面")
            action_a, action_b = visual_value(earlier, "动作/情绪"), visual_value(later, "动作/情绪")
            same_movement = visual_value(earlier, "运镜") == visual_value(later, "运镜")
            pair = f"镜头{shot_index}→{shot_index + 1}"
            if camera_a and picture_a and action_a and camera_a == camera_b and picture_a == picture_b and action_a == action_b and same_movement:
                findings.append(Finding("ERROR", label, f"{pair}的机位、运镜、画面和动作文本相同；按两段台词的语义落点重新选择可辨的新画面"))
            elif camera_a and picture_a and camera_a == camera_b and picture_a == picture_b and same_movement:
                findings.append(Finding("WARN", label, f"{pair}使用相同机位、运镜和画面；核对动作差异是否需要独立切镜"))
            elif (
                same_movement
                and visual_similarity(camera_a, camera_b) >= 0.88
                and visual_similarity(picture_a, picture_b) >= 0.88
                and visual_similarity(action_a, action_b) >= 0.75
            ):
                findings.append(Finding("WARN", label, f"{pair}的画面描述高度相似；逐项核对实际机位、主体、前景和动作"))

        visual_fields = ("摄影机位置", "画面", "动作/情绪")
        first_visual = {field: visual_value(shot_blocks[0], field) for field in visual_fields}
        last_visual = {field: visual_value(shot_blocks[-1], field) for field in visual_fields}
        if previous_boundary:
            _, previous_visual = previous_boundary
            same_camera = bool(first_visual["摄影机位置"]) and first_visual["摄影机位置"] == previous_visual["摄影机位置"]
            same_picture = bool(first_visual["画面"]) and first_visual["画面"] == previous_visual["画面"]
            same_action = bool(first_visual["动作/情绪"]) and first_visual["动作/情绪"] == previous_visual["动作/情绪"]
            if same_camera and same_picture and same_action:
                findings.append(Finding("ERROR", label, "边界两镜的机位、画面和动作文本相同；分别写出两镜具体的主拍人物、机位参照、前景与动作阶段"))
            elif same_camera and same_picture:
                findings.append(Finding("WARN", label, "边界两镜的机位和画面文本相同；核对两幅实际画面的可剪接差异"))
            elif (
                visual_similarity(first_visual["摄影机位置"], previous_visual["摄影机位置"]) >= 0.88
                and visual_similarity(first_visual["画面"], previous_visual["画面"]) >= 0.88
                and visual_similarity(first_visual["动作/情绪"], previous_visual["动作/情绪"]) >= 0.75
            ):
                findings.append(Finding("WARN", label, "边界两镜的机位、画面和动作描述高度相似；逐项比较两幅实际画面"))
        previous_boundary = (label, last_visual)

        if computed > 85:
            findings.append(Finding("ERROR", label, f"实际有声字数{computed}超过85"))
        if declared_match and int(declared_match.group(1)) != computed:
            findings.append(
                Finding("ERROR", label, f"声明有声字数{declared_match.group(1)}与实际{computed}不符")
            )
        if voiced_shots == 0:
            findings.append(Finding("ERROR", label, "整节零有声；与相邻有声剧情合节"))
        elif duration >= 13:
            lower_target = 5 * duration - 5
            upper_target = 5.5 * duration + 2.5
            if computed < lower_target:
                findings.append(Finding("WARN", label, f"有声{computed}字低于普通节约{lower_target:g}字的目标；若保留关键画面，按约50字核对"))
            if computed > upper_target:
                findings.append(Finding("WARN", label, f"有声{computed}字高于本时长普通节约{upper_target:g}字的目标；检查朗读与动作时间"))
        if not re.search(r"^←(?:故事开篇|承上视频编号)", section, re.M):
            findings.append(Finding("ERROR", label, "缺少故事开篇或承上说明"))
        if not re.search(r"^→(?:衔接视频编号|本章结束)", section, re.M):
            findings.append(Finding("ERROR", label, "缺少编号衔接或本章结束说明"))

    if total_shots >= 30 and moving_shots == 0:
        findings.append(Finding("WARN", "全文", f"共{total_shots}镜均未写运镜；回看剧情中的走位、转折和关键信息，选择有动机的少量运镜"))
    if moving_shots >= 5:
        most_common_family, family_count = movement_counts.most_common(1)[0]
        if family_count / moving_shots >= 0.8:
            findings.append(Finding("WARN", "全文", f"{moving_shots}个运镜中{family_count}个为{most_common_family}；核对各镜画面任务，保留有效运动并调整重复设计"))

    return findings


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("draft", type=Path, help="UTF-8 storyboard text file")
    args = parser.parse_args()
    text = args.draft.read_text(encoding="utf-8-sig")
    findings = audit(text)
    for finding in findings:
        print(finding.render())
    errors = sum(finding.level == "ERROR" for finding in findings)
    warnings = sum(finding.level == "WARN" for finding in findings)
    if errors:
        print(f"FAILED: {errors} error(s), {warnings} warning(s)")
        return 1
    print(f"PASSED: 0 errors, {warnings} warning(s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
