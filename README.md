# 進学館烏丸御池校 スケジュール作成アプリ

月次スケジュール表を Word / Excel 形式で生成する Streamlit アプリです。

## 使い方（Web版）

ブラウザで以下を開くだけで使えます。会社のPCからも利用可能です。

**https://karasuma-schedule-8w5houzrnsvuzkf6sco5vu.streamlit.app/**

しばらく使っていないとスリープ状態になります。その場合は画面に出る
「Yes, get this app back up!」を押すと 1 分ほどで復帰します。

## 使い方（ローカル）

```bash
python3 -m streamlit run app.py --server.port 8502
```

ブラウザで http://localhost:8502 を開きます。
Mac ではフォルダ内の「アプリを起動.command」をダブルクリックしても起動できます。

## 入力項目

1. 対象期間（開始日・終了日。夏期講習プリセットあり）
2. 追加イベント（保護者会・特別授業など。年間固定イベントは自動反映）
3. 授業なし期間（夏期・冬期・春期講習。週番号と通常授業を非表示にする）
4. 特別お知らせ（保護者向けの詳細説明文）
5. 出力ファイル名

## ファイル構成

| ファイル | 役割 |
|---|---|
| `app.py` | 画面（入力フォーム） |
| `generate_schedule.py` | Word / Excel 生成処理 |
| `annual_schedule.json` | 週間固定時間割・年間固定イベント |
| `requirements.txt` | 依存ライブラリ |

年間固定イベントや時間割を変更する場合は `annual_schedule.json` を編集してください。

## 備考

イラスト素材（`assets/`・`illustrations/`）は配布ライセンスの都合上リポジトリに含めていません。
生成した Word ファイルに手作業で貼り付けてください。
