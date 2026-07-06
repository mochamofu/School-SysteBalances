# 積立金会計 入力アシスタント（School-SysteBalances）

私立高校の「令和○年度生積立金」（学校徴収金個人別管理票）への入力業務を
自動化・安全化する商品パッケージのリポジトリです。

## リポジトリの地図（どこに何があるか）

```
School-SysteBalances/
├── assistant/                 ★製品本体（これが商品）
│   ├── vba/                   VBAモジュール A00〜A11（12本・ボタン①〜⑮）
│   ├── src/                   生成スクリプト（アシスタント本体・練習/検証データ・USBキット）
│   ├── output/                生成物置き場（git管理外）
│   └── README.md              製品の説明・導入手順・安全装置
│
├── docs/                      資料（シーン別に番号フォルダで分類）
│   ├── 01_学校向けマニュアル/   導入済みの学校に渡す使い方資料
│   ├── 02_営業・商談資料/       提案書・価格比較・スライド・料金設計メモ
│   ├── 03_現地デモ・レクチャー/  訪問デモ（2台PCシミュレーション）の手順書
│   ├── 04_オンライン導入/       郵送USB＋Zoomで導入するときの一式
│   ├── 05_開発・検証記録/       テストログ・実ファイル検証の記録
│   ├── _build/                上記PDF/スライドを生成するPythonスクリプト
│   ├── animations/            操作アニメーションGIFと生成スクリプト
│   └── README.md              docs内の詳しい案内
│
├── site/                      公開用ランディングページ（商品LP・料金表つき）
│
├── archive_第1世代ツール/      実物ファイル入手前に作った旧ツール一式（参考保存）
│
└── .github/workflows/         LPのGitHub Pages公開（手動実行のみ）
```

## シーン別の入口

| やりたいこと | 見る場所 |
|---|---|
| 製品そのものを知る・作る | `assistant/README.md` → `python assistant/src/build_assistant.py` |
| 学校へ営業に行く | `docs/02_営業・商談資料/`（提案書・価格1枚・スライド） |
| 現地でデモ・レクチャーする | `docs/03_現地デモ・レクチャー/` ＋ デモキット生成 `python assistant/src/make_usb_kit.py` |
| 郵送＋Zoomで導入する（全国展開） | `docs/04_オンライン導入/` ＋ 導入USB生成 `python assistant/src/make_school_usb_kit.py` |
| 導入済みの学校をサポートする | `docs/01_学校向けマニュアル/` |
| 品質の根拠を確認する | `docs/05_開発・検証記録/` |
| LPを公開・更新する | `site/`（Netlify Drop / Cloudflare Pages 推奨。詳細は `site/README.md`） |

## 生成物（キット）の作り方

```bash
pip install openpyxl python-docx python-pptx pillow

python assistant/src/build_assistant.py        # アシスタント本体（xlsx）
python assistant/src/make_practice_files.py    # 練習用データ（架空80名）
python assistant/src/make_fullscale_test_files.py  # 検証用データ（架空320名）
python assistant/src/make_usb_kit.py           # 営業・デモ用キット（自分が持ち歩く）
python assistant/src/make_school_usb_kit.py    # 学校郵送用USB 2本（送りっぱなし用）
```

いずれも `assistant/output/` に出力されます（git管理外）。
PDF・スライドを作り直す場合は `docs/_build/` の各スクリプトを実行してください。

## だいじな約束

- 実在の生徒・口座データはリポジトリにもキットにも一切入れない（すべて架空データ）
- 学校のマスターファイルには行も列も追加しない。書き込み前は必ず自動バックアップ
