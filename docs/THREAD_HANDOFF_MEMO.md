# スレ引っ越しメモ

## プロジェクト名

料金改定計算DB 改良版 v0.1

## リポジトリ

Gootarou/price-revision-db-v01

## 技術構成

* Googleスプレッドシート
* Google Apps Script
* clasp
* GitHub
* Codex

## ローカル作業フォルダ

`C:\Users\syama\20_output\00_開発\price-revision-db-v01`

Apps Script フォルダ：

`C:\Users\syama\20_output\00_開発\price-revision-db-v01\apps-script`

## 重要な運用ルール

* `.clasp.json` はローカルのみです。
* `.clasp.json` はGit管理しません。
* `clasp create` は通常使いません。
* 現在は `.clasp.json` が指しているスプレッドシートを正とします。
* PRをマージしたら `git pull origin main` します。
* GAS反映は `apps-script` フォルダで `clasp push` します。
* スプレッドシートを再読み込みしてメニュー確認します。

## これまでの実装フェーズ

* Phase 1：初期セットアップ
* Phase 2：入力画面
* Phase 3：保存・読込
* Phase 4：月額契約計算
* Phase 5：勤務時間根拠確認サイドバー
* Phase 6：文書転記前確認ビュー
* Phase 7：README・テスト・運用メモ
* Phase 8：v0.1凍結前チェック・引っ越しメモ

## 現在の到達点

v0.1の主要機能は実装済みです。
次は新機能追加ではなく、手動テストと実運用テストを行う段階です。

## 次スレで最初にやること

1. PR #12 と Phase 8 のPRがmainへマージ済みか確認する。
2. ローカルで `git pull origin main` する。
3. `apps-script` で `clasp push` する。
4. スプレッドシートを再読み込みする。
5. `tests/manual-test-checklist.md` に沿って実機テストする。
6. テスト結果から修正要否を判断する。

## 次スレで優先して確認すること

* 入力画面初期化
* 保存
* 読込
* 保存して計算
* 勤務時間根拠サイドバー
* 文書転記前確認ビュー
* エラー系
* 月額以外の対象外処理

## 次スレで安易にやらないこと

* 新機能追加
* 文書生成
* PDF生成
* AppSheet化
* onEdit追加
* セル保護
* 条件付き書式自動生成
* 未保存警告
* 保存前確認ダイアログ

これらは実運用テスト後に、必要性を確認してから検討します。
