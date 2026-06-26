"""ローカル CLI: 旅行命令復命書 Excel → 動静表 Excel.

使い方:
    python 動静表生成.py [入力.xlsx] [出力.xlsx]

引数を省略した場合は、同じフォルダの「旅行命令復命書.xlsx」を読み込み、
入力データの期間から自動命名したファイルを出力します。
"""

import sys
from pathlib import Path

import processor


def main():
    script_dir = Path(__file__).parent
    input_path = Path(sys.argv[1]) if len(sys.argv) >= 2 else script_dir / "旅行命令復命書.xlsx"

    if not input_path.exists():
        print(f"入力ファイルが見つかりません: {input_path}")
        sys.exit(1)

    with open(input_path, "rb") as f:
        out_bytes, out_name, days, entries = processor.process_excel(f)

    output_path = Path(sys.argv[2]) if len(sys.argv) >= 3 else script_dir / out_name
    output_path.write_bytes(out_bytes)
    print(f"出力完了: {output_path}  ({days}日分 / {entries}件)")


if __name__ == "__main__":
    main()
