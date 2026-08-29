#!/usr/bin/env python3
"""gakkai-slides check_deck.py: 学会発表pptxをデザイン規則に照らして検査する。

使い方:
    python3 check_deck.py deck.pptx
    python3 check_deck.py deck.pptx --json

見るもの:
    塗りつぶし図形（box）、角丸カード、ページ番号の欠落、フォント（Noto Sans JP以外）、
    パレット外の色、赤の使いすぎ、16pt未満の本文、テキスト枠の乱立（1文1ボックスの疑い）、
    1枚あたりの行数過多。検出は疑いの提示であって、直すかどうかは人が決める。
依存: python-pptx
"""
from __future__ import annotations

import argparse
import json
import re
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt
from pptx.oxml.ns import qn

findings = []


def add(level, slide_no, check, message):
    findings.append({"level": level, "slide": slide_no, "check": check, "message": message})


def rgb_of(font):
    try:
        if font.color and font.color.type is not None and font.color.rgb is not None:
            return str(font.color.rgb)
    except Exception:
        pass
    return None


def is_allowed_color(hexstr):
    r, g, b = int(hexstr[0:2], 16), int(hexstr[2:4], 16), int(hexstr[4:6], 16)
    if abs(r - g) <= 16 and abs(g - b) <= 16:      # 黒〜グレー〜白
        return True
    if r >= 0xA0 and g <= 0x60 and b <= 0x60:      # 赤系
        return True
    return False


def is_red(hexstr):
    r, g, b = int(hexstr[0:2], 16), int(hexstr[2:4], 16), int(hexstr[4:6], 16)
    return r >= 0xA0 and g <= 0x60 and b <= 0x60


def ea_font(run):
    rPr = run._r.find(qn("a:rPr"))
    if rPr is None:
        return None
    ea = rPr.find(qn("a:ea"))
    return ea.get("typeface") if ea is not None else None


def check(path, font_name):
    prs = Presentation(path)
    n_slides = len(prs.slides._sldIdLst)
    pagenum_re = re.compile(r"^\s*\d+\s*(/\s*\d+)?\s*$")

    for idx, slide in enumerate(prs.slides, 1):
        frames = 0
        red_runs = 0
        lines = 0
        has_pagenum = False

        for shape in slide.shapes:
            # 図形（box）の検査
            if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                add("警告", idx, "図形box", f"オートシェイプ「{shape.shape_type}」がある。boxは使わない")
            if shape.shape_type in (MSO_SHAPE_TYPE.TEXT_BOX, MSO_SHAPE_TYPE.AUTO_SHAPE):
                try:
                    if shape.fill.type is not None and shape.fill.type == 1:  # solid
                        hexstr = str(shape.fill.fore_color.rgb)
                        if hexstr.upper() != "FFFFFF":
                            add("警告", idx, "塗りつぶし", f"塗り #{hexstr} の枠がある。白背景に文字だけにする")
                except Exception:
                    pass
            try:
                if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and "ROUND" in shape.auto_shape_type.name:
                    add("警告", idx, "角丸カード", "角丸図形はAI生成スライドの典型。使わない")
            except Exception:
                pass

            if not shape.has_text_frame:
                continue
            text_all = shape.text_frame.text.strip()
            if text_all:
                frames += 1
            if pagenum_re.match(text_all):
                has_pagenum = True

            for para in shape.text_frame.paragraphs:
                if para.text.strip():
                    lines += 1
                for run in para.runs:
                    if not run.text.strip():
                        continue
                    # フォント
                    for name in (run.font.name, ea_font(run)):
                        if name and not name.startswith(font_name):
                            add("情報", idx, "フォント", f"「{name}」（既定は {font_name}）: {run.text[:14]}")
                            break
                    # サイズ
                    if run.font.size is not None:
                        pt = run.font.size.pt
                        if pt < 12:
                            add("警告", idx, "文字サイズ", f"{pt:.0f}pt（12pt未満）: {run.text[:14]}")
                        elif pt < 16 and not pagenum_re.match(text_all):
                            hexstr0 = rgb_of(run.font)
                            grayish = hexstr0 and abs(int(hexstr0[0:2],16)-int(hexstr0[2:4],16))<=16 and 0x30<=int(hexstr0[0:2],16)<=0x90
                            if not grayish:
                                add("情報", idx, "文字サイズ", f"{pt:.0f}pt。本文は16pt以上、12〜14ptはグレーの出典のみ: {run.text[:14]}")
                    # 色
                    hexstr = rgb_of(run.font)
                    if hexstr:
                        if not is_allowed_color(hexstr):
                            add("警告", idx, "パレット外の色", f"#{hexstr}: {run.text[:14]}。黒・赤・グレーに収める")
                        if is_red(hexstr):
                            red_runs += 1

        if idx >= 2 and not has_pagenum:
            add("警告", idx, "ページ番号", "右下に n / N がない")
        if red_runs > 2:
            add("情報", idx, "赤の使いすぎ", f"赤のrunが{red_runs}箇所。強調は1枚2箇所まで")
        if frames > 6:
            add("情報", idx, "テキスト枠の乱立", f"{frames}枠。1文ごとにboxを分けず、本文は1枠にまとめる")
        if lines > 12:
            add("情報", idx, "行数過多", f"{lines}行。話すことだけ書き、残りは補足スライドへ")

    return n_slides


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--font", default="Noto Sans")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    n = check(args.pptx, args.font)
    order = {"重大": 0, "警告": 1, "情報": 2}
    findings.sort(key=lambda f: (order[f["level"]], f["slide"]))

    if args.json:
        print(json.dumps({"slides": n, "findings": findings}, ensure_ascii=False, indent=2))
        return
    counts = {k: sum(1 for f in findings if f["level"] == k) for k in order}
    print(f"== {args.pptx} == 全{n}枚 / 重大 {counts['重大']} / 警告 {counts['警告']} / 情報 {counts['情報']}")
    for f in findings:
        print(f"[{f['level']}] p{f['slide']} {f['check']}: {f['message']}")
    if not findings:
        print("検出なし。")


if __name__ == "__main__":
    main()
