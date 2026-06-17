# AUS_Report 作成発注書

---

## アプリ概要・目的

院内AUS（人工妊娠中絶）報告書管理システムの再実装版。
現行 `AUS_app` と同じ業務フローを維持しつつ、データ管理を SQLite に移行し、
PowerPoint 生成ロジックを整理してメンテナンス性を向上させる。

- **提出様式（PPTX テンプレート）は変更不可**。現行アプリで使用中の正式様式をそのまま使用する。
- 現行アプリ（ポート 50003）と並行稼働できるよう、ポートは **50004** とする。

---

## 機能一覧

| # | 機能 | 詳細 |
|---|---|---|
| 1 | データ登録 | 通算No・年齢・管轄・週数・日付・ID・特記事項を入力して保存 |
| 2 | データ表示 | 月別フィルタ・件数表示・ソート（日付順/ID順/通算No順） |
| 3 | データ編集 | 同一IDで日付違いのレコードを正確に1件だけ編集 |
| 4 | データ削除 | ID＋日付の組み合わせで1件のみ削除・削除ログ記録 |
| 5 | レポート生成 | 月次 Excel（月別データ）＋ PowerPoint 2種（報告書・実施数報告）を生成 |
| 6 | エラー画面 | エラー時に `error.html` を表示しトップへ戻れる |

---

## 画面構成・UI 仕様

- **デザイン**: 現行アプリ（ダークテーマ、Bootstrap 5、メイリオ）を踏襲
- **index.html**: 左カラム＝データ入力フォーム、右カラム＝月別データ一覧＋レポート生成
- **edit.html**: 編集フォーム（キャンセルボタンでトップへ戻る）
- **error.html**: エラーメッセージ＋トップへ戻るリンク
- Bootstrap は CDN を使わず `static/css/` `static/js/` にローカル配置（オフラインLAN対応）
- フラッシュメッセージは 8 秒表示後フェードアウト

---

## データ構造（SQLite）

### テーブル: `cases`

```sql
CREATE TABLE IF NOT EXISTS cases (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    seq_no           INTEGER NOT NULL,          -- 通算No
    age              INTEGER NOT NULL,
    address          TEXT    NOT NULL,          -- 管轄（住所文字列）
    gestational_age  INTEGER NOT NULL,          -- 週数（5〜21 の生値）
    date             TEXT    NOT NULL,          -- YYYY-MM-DD
    patient_id       TEXT    NOT NULL,          -- 5桁ゼロ埋め文字列
    note             TEXT    NOT NULL DEFAULT '',  -- 特記事項
    UNIQUE(patient_id, date)
);
```

**補足:**
- `gestational_age` は入力時の生の週数（5〜21）を保存する。レポート生成時に変換する。
- `patient_id` は `zfill(5)` で 5 桁にゼロ埋めして保存する。

---

## ファイル構成

```
AUS_Report/
├── app.py                      # Flask アプリ本体
├── config.py                   # 設定クラス
├── requirements.txt            # 依存パッケージ
├── .gitignore
├── .env.example
├── start_app.bat               # 起動バッチ
├── AUS報告書テンプレ.pptx       # ※ 提出用正式様式（変更禁止）
├── AUS実施数報告テンプレ.pptx   # ※ 提出用正式様式（変更禁止）
├── templates/
│   ├── index.html
│   ├── edit.html
│   └── error.html
├── static/
│   ├── css/bootstrap.min.css
│   └── js/bootstrap.bundle.min.js
└── docs/
    ├── 修正指示書.md
    └── 修正実行書.md
```

---

## 設定クラス（config.py）

```python
import os

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'aus-report-change-in-production'
    DB_PATH = 'aus_data.db'
    TEMPLATE_PATH = 'AUS報告書テンプレ.pptx'
    IMPLEMENTATION_TEMPLATE_PATH = 'AUS実施数報告テンプレ.pptx'
    OUTPUT_DIR = 'month_data'
    PPTX_OUTPUT_DIR = 'AUS報告書'
    HOST = '0.0.0.0'
    PORT = 50004
    DEBUG = False
```

---

## セットアップ手順

```bat
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py   ← 初回起動時に DB を自動作成
```

### requirements.txt

```
Flask==3.1.2
pandas>=2.0.0
openpyxl>=3.1.0
python-pptx>=0.6.21
python-dotenv>=1.0.0
```

### .gitignore（最低限）

```
venv/
__pycache__/
*.pyc
.env
*.db
instance/
.DS_Store
changes.log
*.log
logs/
month_data/
AUS報告書/
*.xlsx
```

### start_app.bat

```bat
@echo off
cd /d "%~dp0"
call "%~dp0venv\Scripts\activate.bat"
python "%~dp0app.py"
```

---

## app.py 実装仕様

### DB 初期化

アプリ起動時（`if __name__ == '__main__'` の前）に `init_db()` を呼び、
テーブルが存在しない場合のみ作成する。

```python
def get_db():
    conn = sqlite3.connect(Config.DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as conn:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS cases (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                seq_no          INTEGER NOT NULL,
                age             INTEGER NOT NULL,
                address         TEXT    NOT NULL,
                gestational_age INTEGER NOT NULL,
                date            TEXT    NOT NULL,
                patient_id      TEXT    NOT NULL,
                note            TEXT    NOT NULL DEFAULT '',
                UNIQUE(patient_id, date)
            )
        ''')
```

### ルート一覧

| メソッド | URL | 処理 |
|---|---|---|
| GET | `/` | 月別一覧表示（月パラメータ: `?month=YYYY-MM`） |
| POST | `/submit` | 新規登録 |
| GET | `/edit/<patient_id>/<date>` | 編集フォーム表示 |
| POST | `/submit_edit/<patient_id>/<date>` | 編集保存 |
| POST | `/delete/<patient_id>/<date>` | 削除 |
| POST | `/generate_report` | レポート生成 |

### バリデーション（共通）

- 全必須フィールドの存在確認
- `age`: 0〜100
- `gestational_age`: 0〜45
- `date`: `YYYY-MM-DD` 形式・未来日付不可
- `patient_id`: 5桁数字
- 同一 `patient_id` ＋ `date` の重複チェック（編集時は自分自身を除外）

### 削除ログ（changes.log）

削除時に `changes.log` へ追記する（`.gitignore` で除外済み）。

```
YYYY-MM-DD HH:MM:SS - 削除
seq_no: X
patient_id: XXXXX
date: YYYY-MM-DD
...
```

---

## PowerPoint 生成仕様

### 重要: テンプレートのプレースホルダー名は変更禁止

正式様式テンプレートに埋め込まれたテキストボックスのプレースホルダー文字列は
以下のとおり。**1文字も変えてはならない。**

| プレースホルダー | 内容 | フォントサイズ |
|---|---|---|
| `RY{n}` | 令和年 | 10pt |
| `MM{n}` | 報告月 | 10pt |
| `N{n}` | 症例番号（月内連番） | 12pt |
| `A{n}` | 年齢 | 12pt |
| `AD{n}-1` | 都道府県名（「県」等を除く） | 12pt |
| `AD{n}-2` | 市/郡の地名 | 12pt |
| `AD{n}-3` | 町/村名 | 12pt |
| `M{n}` | 症例日付の月 | 12pt |
| `D{n}` | 症例日付の日 | 12pt |
| `G{n}-1` | 週数区分1（5〜7週）の〇 | 14pt |
| `G{n}-2` | 週数区分2（8〜11週）の〇 | 14pt |
| `GUN{n}` | 郡の〇 | 14pt |
| `SHI{n}` | 市の〇 | 14pt |
| `MACHI{n}` | 町の〇 | 14pt |

`{n}` は症例番号（1 または 2、スライド1枚に2症例）。

### 週数の変換（レポート生成時のみ）

| 生の週数 | 変換値 |
|---|---|
| 5〜7週 | `'1'` → G{n}-1 に〇 |
| 8〜11週 | `'2'` → G{n}-2 に〇 |
| 12〜21週、その他 | `''` → G{n}-1・G{n}-2 ともに空 |

### 住所の分解

```python
def parse_address(address: str) -> tuple[str, str, str, str]:
    """
    住所文字列を (都道府県, 市/郡地名, '市'/'郡', 町/村名) に分解する。
    例: '○○県○○市○○町' → ('○○', '○○', '市', '○○')
    """
```

正規表現で都道府県→市/郡→町/村の順に抽出する。
現行 AUS_app の `parse_address_for_excel()` と同じロジックを実装すること。

### PPT 生成の実装方針（クリーン版）

**スライド1枚に症例を最大2件配置する**（`case_num` = 1 or 2）。

```python
def build_replacements(row, case_num, reiwa_year, report_month):
    """症例1件分のプレースホルダー→置換値 辞書を返す"""
    n = str(case_num)
    date_obj = pd.to_datetime(row['date'])
    ga_raw = int(row['gestational_age'])
    ga = convert_gestational_age(ga_raw)   # '1','2', or ''
    pref, city, gunshi, town = parse_address(row['address'])

    return {
        f'RY{n}':    (str(reiwa_year),     10),
        f'MM{n}':    (str(report_month),   10),
        f'N{n}':     (str(row['number']),  12),
        f'A{n}':     (str(row['age']),     12),
        f'AD{n}-1':  (pref,                12),
        f'AD{n}-2':  (city,                12),
        f'AD{n}-3':  (town,                12),
        f'M{n}':     (str(date_obj.month), 12),
        f'D{n}':     (str(date_obj.day),   12),
        f'G{n}-1':   ('〇' if ga == '1' else '', 14),
        f'G{n}-2':   ('〇' if ga == '2' else '', 14),
        f'GUN{n}':   ('〇' if gunshi == '郡' else '', 14),
        f'SHI{n}':   ('〇' if gunshi == '市' else '', 14),
        f'MACHI{n}': ('〇' if town else '', 14),
    }


def apply_replacements(slide, replacements):
    """
    スライド内の全シェイプ・全段落に対して1パスで置換を行う。
    キーの長いものから順に置換し、部分一致を防ぐ。
    """
    sorted_keys = sorted(replacements.keys(), key=len, reverse=True)
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for key in sorted_keys:
                if key in para.text:
                    value, font_pt = replacements[key]
                    new_text = para.text.replace(key, value)
                    para.clear()
                    run = para.add_run()
                    run.text = new_text
                    run.font.size = Pt(font_pt)
                    break  # 1段落につき1置換
```

### 実施数報告テンプレートのプレースホルダー

| プレースホルダー | 内容 |
|---|---|
| `G1` | 令和年 |
| `H1` | 報告月 |
| `I1` | レポート生成日の月 |
| `J1` | レポート生成日の日 |
| `K1` | 月内最終症例番号 |

### 不要なスライドの削除

データ件数が奇数の場合、最終スライドの2件目は空のまま残る。
`required_slides = (len(data) + 1) // 2` を超えるスライドは末尾から削除する。

---

## セキュリティ要件（院内ネットワーク限定）

- `SECRET_KEY` は `.env` で管理（コードにハードコードしない）
- `app.run(debug=False)` を本番運用時に設定
- SQL はパラメータバインディングのみ使用（f-string での SQL 組み立て禁止）
- テンプレートでは `{{ }}` の自動エスケープを維持（`| safe` 禁止）
- `HOST = '0.0.0.0'` でLAN公開するため、Windows ファイアウォールで TCP 50004 を許可すること

---

## .cursor/rules

```
# Cursor ルール（AUS_Report）

## 基本方針
- 本発注書の内容に従い忠実にコーディングを行う
- DB スキーマ変更を伴う場合は作業前にユーザーに確認する
- コーディング完了後、docs/修正実行書.md の末尾に実施内容を追記する（新規ファイル作成禁止）
- 修正対象ファイルを必ず読んでから変更する
- SQL はパラメータバインディングを使用（f-string での SQL 組み立て禁止）
- SECRET_KEY をコードにハードコードしない
- AUS報告書テンプレ.pptx・AUS実施数報告テンプレ.pptx の内容は変更禁止
- Claude Code のレビューが完了するまで git commit しない

## メインファイル
- app.py（Flask アプリ本体）

## ドキュメント
- docs/修正指示書.md（修正指示）
- docs/修正実行書.md（実行記録）
```

---

## タスクスケジューラ登録

- トリガー　: コンピューターの起動時
- 操作　　　: `start_app.bat` を実行
- 全般設定　: 最上位の特権で実行・ユーザーがログオンしていなくても実行

---

## 注意事項

1. **テンプレートファイルは変更禁止**: `AUS報告書テンプレ.pptx` と `AUS実施数報告テンプレ.pptx` はプロジェクトルートに配置するだけで、内容を変えてはならない。
2. **プレースホルダー名は完全一致**: テンプレート内のプレースホルダー文字列（`RY1`, `G1-1`, `SHI1` 等）は1文字も変えてはならない。
3. **部分一致を防ぐ置換順序**: `D{n}` より `AD{n}-1` を先に置換すること（`apply_replacements` 内で長いキーから順に処理）。
4. **Bootstrap はローカル配置**: `static/css/bootstrap.min.css` と `static/js/bootstrap.bundle.min.js` を配置し CDN を使わないこと。
5. **初回起動**: `init_db()` をアプリ起動時に呼び、DB が存在しない場合のみテーブルを作成する。
6. **現行アプリとの並行稼働**: ポート 50004 を使用し、現行アプリ（50003）と競合しないこと。
```
