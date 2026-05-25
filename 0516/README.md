# 0516 課題フォルダ（ISM・DEMATEL・因果ループ）

## 提出ファイル

| ファイル | 用途 |
|----------|------|
| `20260516_構造化レポート.pptx` | **提出用 PowerPoint** |
| `20260516_構造化レポート.pdf` | **提出用 PDF** |
| `20260516_answer.md` | 全文テキスト（編集用） |

## ツール

| ファイル | 用途 |
|----------|------|
| `ism_dematel.html` | ブラウザで ISM/DEMATEL 計算（xlsm 代替） |
| `compute_ism_dematel.py` | Python で同計算 → `results.json` |
| `generate_presentation.py` | pptx 再生成 |

## xlsm への入力

配布 `2026_ISM-DEMATEL.xlsm` に手入力する場合は `results.json` の `A`（ISM）と `X`（DEMATEL）を参照。

問題テーマを変えるときは `compute_ism_dematel.py` の `ITEMS`, `A`, `X` を編集してから再実行。
