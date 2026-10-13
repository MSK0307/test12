"""料理メニューを Claude で分析し、Excel (menu_db.xlsx) に追記するツール。

使い方:
    python analyze_menu.py 肉じゃが 鯖の味噌煮
    python analyze_menu.py -f menus.txt        # 1行に1メニュー
    python analyze_menu.py -f menus.txt --force  # 登録済みのメニューも再分析する
"""

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

import anthropic
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

MODEL = "claude-opus-5-5"
DB_PATH = Path(__file__).with_name("menu_db.xlsx")
SHEET_NAME = "メニュー分析"

INGREDIENT_CATEGORIES = [
    "肉", "魚介", "野菜", "いも", "きのこ", "海藻", "豆・大豆製品",
    "卵", "乳製品", "穀物", "果物", "その他",
]
DISH_FORMS = [
    "汁物", "煮物", "焼き物", "炒め物", "揚げ物", "蒸し物", "和え物・サラダ",
    "生もの・刺身", "鍋物", "ご飯もの", "麺類", "パン・粉もの", "デザート", "その他",
]
SHAPES = ["塊・大きめ", "一口大", "薄切り・細切り", "刻み・ほぐし", "ペースト・ムース状", "液状・汁気多め"]
SOFTNESS_LABELS = {1: "とても柔らかい", 2: "柔らかい", 3: "普通", 4: "やや硬い", 5: "硬い"}
BASIC_TASTES = ["甘味", "塩味", "酸味", "苦味", "うま味", "辛味"]
STRENGTHS = ["薄味", "普通", "濃いめ"]

# Excel の列（見出し, 分析結果から値を取り出す関数）
COLUMNS = [
    ("メニュー名", lambda name, r: name),
    ("主な食材", lambda name, r: "、".join(r["ingredients"])),
    ("食材カテゴリ", lambda name, r: "、".join(r["ingredient_categories"])),
    ("料理形態", lambda name, r: r["dish_form"]),
    ("形状", lambda name, r: r["shape"]),
    ("柔らかさ(1-5)", lambda name, r: r["softness"]),
    ("柔らかさ", lambda name, r: SOFTNESS_LABELS[r["softness"]]),
    ("基本味", lambda name, r: "、".join(r["basic_tastes"])),
    ("味付け", lambda name, r: "、".join(r["seasonings"])),
    ("味の濃さ", lambda name, r: r["taste_strength"]),
    ("補足", lambda name, r: r["notes"]),
    ("分析日時", lambda name, r: datetime.now().strftime("%Y-%m-%d %H:%M")),
]

SCHEMA = {
    "type": "object",
    "properties": {
        "ingredients": {
            "type": "array", "items": {"type": "string"},
            "description": "一般的なレシピで使われる主な食材（調味料は除く）",
        },
        "ingredient_categories": {
            "type": "array", "items": {"type": "string", "enum": INGREDIENT_CATEGORIES},
        },
        "dish_form": {"type": "string", "enum": DISH_FORMS},
        "shape": {"type": "string", "enum": SHAPES},
        "softness": {"type": "integer", "enum": list(SOFTNESS_LABELS)},
        "basic_tastes": {
            "type": "array", "items": {"type": "string", "enum": BASIC_TASTES},
            "description": "はっきり感じられる基本味。強いものから順に",
        },
        "seasonings": {
            "type": "array", "items": {"type": "string"},
            "description": "主な味付け（醤油、味噌、トマトソースなど）",
        },
        "taste_strength": {"type": "string", "enum": STRENGTHS},
        "notes": {"type": "string", "description": "判定の前提や、作り方で変わる点を一言で"},
    },
    "required": [
        "ingredients", "ingredient_categories", "dish_form", "shape", "softness",
        "basic_tastes", "seasonings", "taste_strength", "notes",
    ],
    "additionalProperties": False,
}

SYSTEM_PROMPT = """あなたは料理の分析担当です。メニュー名から、日本の家庭や飲食店で一般的に作られる形を想定して、その料理を分析してください。

- 食材: 主な食材を挙げ、該当する食材カテゴリをすべて選ぶ。
- 形態: 調理法から料理形態を1つ、料理に入っている主な具の形状を1つ選ぶ。
- 柔らかさ: 料理全体の食感を 1(とても柔らかい: 舌でつぶせる) 〜 5(硬い: よく噛む必要がある) で評価する。
- 味: はっきり感じられる基本味、主な味付け、味の濃さを答える。
- メニュー名があいまいな場合は、最も一般的な解釈で答え、その前提を notes に書く。"""


def analyze(client: anthropic.Anthropic, menu_name: str) -> dict:
    response = client.beta.messages.create(
        model=MODEL,
        max_tokens=8000,
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",
        output_config={
            "effort": "low",
            "format": {"type": "json_schema", "schema": SCHEMA},
        },
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": f"メニュー名: {menu_name}"}],
    )
    if response.stop_reason == "refusal":
        raise RuntimeError("Claude が分析を断りました")
    if response.stop_reason == "max_tokens":
        raise RuntimeError("出力が上限に達して途中で切れました")
    text = next(b.text for b in response.content if b.type == "text")
    return json.loads(text)


def open_db() -> tuple[Workbook, object]:
    if DB_PATH.exists():
        wb = load_workbook(DB_PATH)
        return wb, wb[SHEET_NAME]
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append([header for header, _ in COLUMNS])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4F6D7A")
        cell.alignment = Alignment(horizontal="center")
    widths = [18, 30, 20, 12, 16, 12, 14, 18, 20, 10, 40, 17]
    for i, width in enumerate(widths):
        ws.column_dimensions[chr(ord("A") + i)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return wb, ws


def find_row(ws, menu_name: str) -> int | None:
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == menu_name:
            return row
    return None


def main() -> int:
    parser = argparse.ArgumentParser(description="料理メニューを分析して menu_db.xlsx に保存します")
    parser.add_argument("menus", nargs="*", help="メニュー名（複数可）")
    parser.add_argument("-f", "--file", type=Path, help="メニュー名を1行に1つ書いたテキストファイル")
    parser.add_argument("--force", action="store_true", help="登録済みのメニューも再分析して上書きする")
    args = parser.parse_args()

    menus = list(args.menus)
    if args.file:
        menus += [line.strip() for line in args.file.read_text(encoding="utf-8").splitlines()]
    menus = list(dict.fromkeys(m for m in menus if m))  # 空行と重複を除く
    if not menus:
        parser.error("メニュー名を指定するか、-f でファイルを指定してください")

    client = anthropic.Anthropic()
    wb, ws = open_db()
    added = failed = skipped = 0

    for name in menus:
        existing = find_row(ws, name)
        if existing and not args.force:
            print(f"- {name}: 登録済みのためスキップ（再分析は --force）")
            skipped += 1
            continue
        try:
            result = analyze(client, name)
        except anthropic.APIConnectionError as e:
            print(f"× {name}: API に接続できません ({e})", file=sys.stderr)
            failed += 1
            continue
        except anthropic.APIStatusError as e:
            print(f"× {name}: API エラー {e.status_code} ({e.message})", file=sys.stderr)
            if isinstance(e, anthropic.AuthenticationError):
                break  # 認証エラーは以降も必ず失敗するので打ち切る
            failed += 1
            continue
        except (RuntimeError, json.JSONDecodeError) as e:
            print(f"× {name}: {e}", file=sys.stderr)
            failed += 1
            continue

        values = [get(name, result) for _, get in COLUMNS]
        if existing:
            for col, value in enumerate(values, start=1):
                ws.cell(row=existing, column=col, value=value)
        else:
            ws.append(values)
        wb.save(DB_PATH)  # 1件ごとに保存し、途中で止まっても結果が残るようにする
        added += 1
        print(f"○ {name}: {result['dish_form']} / 柔らかさ {result['softness']} / "
              f"{'・'.join(result['basic_tastes'])} / {'、'.join(result['ingredients'][:4])}")

    ws.auto_filter.ref = ws.dimensions
    wb.save(DB_PATH)
    print(f"\n完了: 保存 {added} 件、スキップ {skipped} 件、失敗 {failed} 件 → {DB_PATH.name}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
