# 動静表ジェネレーター

旅行命令復命書の Excel ファイルをアップロードすると、日付ごとに集約した動静表 Excel をダウンロードできる Web アプリです。

## 抽出仕様（各シート）

| 項目 | セル |
| --- | --- |
| 氏名 | `W3` |
| 出張開始 | `J4`(令和年) / `M4`(月) / `P4`(日) / `S4`(時) / `V4`(分) |
| 出張終了 | `J5`(令和年) / `M5`(月) / `P5`(日) / `S5`(時) / `V5`(分) |
| 出張先 | `F7` |
| 用務 | `F8` |

出力は A 列に日付、B 列に `氏名：用務（出張先）時分～時分`。同日の予定は同セルに改行で連結します。
出力ファイル名は入力データの期間から自動生成されます（例：`動静表_令和8年6月2日～30日.xlsx`）。

## ローカル実行

```bash
pip install -r requirements.txt
python app.py
# ブラウザで http://localhost:5000 を開く
```

CLI から使う場合：

```bash
python 動静表生成.py [入力.xlsx] [出力.xlsx]
```

## GitHub → Render デプロイ手順

### 1. GitHub にリポジトリを作成して push

ローカルの本フォルダで：

```bash
git init
git add .
git commit -m "初期コミット: 動静表ジェネレーター"
git branch -M main
git remote add origin https://github.com/<あなたのユーザー名>/<リポジトリ名>.git
git push -u origin main
```

> 注: `.gitignore` で `*.xlsx` を除外しているので、個人情報を含む入力サンプルは push されません。

### 2. Render で Web Service を作成

1. https://dashboard.render.com にログイン
2. 「New +」→「Web Service」
3. 上で push した GitHub リポジトリを選択
4. 設定は `render.yaml` が自動で読み込まれます（Blueprint）
   - 手動で入力する場合：
     - Runtime: **Python**
     - Build Command: `pip install -r requirements.txt`
     - Start Command: `gunicorn app:app --bind 0.0.0.0:$PORT`
     - Plan: **Free**
5. 「Create Web Service」をクリック
6. ビルドが終わると `https://<service-name>.onrender.com` で公開されます

### 3. 更新

ローカルでコードを変更したら：

```bash
git add .
git commit -m "更新内容"
git push
```

Render が自動で再デプロイします。

## ファイル構成

```
├── app.py              # Flask Web アプリ
├── processor.py        # 抽出 / 集計 / 出力ロジック
├── 動静表生成.py        # ローカル CLI
├── templates/
│   └── index.html      # アップロード UI
├── requirements.txt
├── render.yaml
├── .gitignore
└── README.md
```
