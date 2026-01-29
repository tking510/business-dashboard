# ダッシュボード自動更新ガイド

## 概要

ExcelファイルをGitHubにアップロードするだけで、自動的にdata.jsが生成され、ダッシュボードが更新されます。

## 自動化の仕組み

1. `data/`ディレクトリにExcelファイルをアップロード
2. GitHubにプッシュ
3. **GitHub Actions**が自動的に起動
4. Pythonスクリプトがdata.jsを生成
5. 自動的にコミット＆プッシュ
6. GitHub Pagesが更新され、ダッシュボードに反映

## 使い方

### 方法1: GitHubウェブインターフェース（簡単）

1. https://github.com/tking510/business-dashboard/tree/master/data にアクセス
2. 「Add file」→「Upload files」をクリック
3. 最新のExcelファイルをドラッグ&ドロップ
4. 「Commit changes」をクリック
5. **数分待つ**とdata.jsが自動更新されます

### 方法2: Git コマンドライン

```bash
cd /path/to/business-dashboard
cp /path/to/latest/files/*.xlsx data/
git add data/
git commit -m "Update data files"
git push
```

## 対応ファイル

以下のファイル名パターンに一致するファイルが自動処理されます：

- **SAMファイル**: `*SAM*.xlsx` または `収益管理*.xlsx`
- **スロ天KPIシート**: `*スロ天*.xlsx` または `*KPI*.xlsx`
- **Konibetデータ**: `*日报*.csv` または `*データ総和*.csv`

## 確認方法

1. GitHubリポジトリの「Actions」タブで実行状況を確認
2. 緑色のチェックマークが表示されたら成功
3. ダッシュボード (https://tking510.github.io/business-dashboard/) を開いて確認
4. **Ctrl+Shift+R**で強制リロードして最新データを表示

## トラブルシューティング

### GitHub Actionsが実行されない

- `.github/workflows/update-data.yml`が正しくプッシュされているか確認
- GitHubリポジトリの「Settings」→「Actions」→「General」で、Actionsが有効になっているか確認

### data.jsが更新されない

- 「Actions」タブでエラーメッセージを確認
- ファイル名が対応パターンに一致しているか確認
- Excelファイルが破損していないか確認

### ダッシュボードに反映されない

- ブラウザのキャッシュをクリア（Ctrl+Shift+R）
- GitHub Pagesの更新に5-10分かかることがあります

## 注意事項

- 複数のファイルを同時にアップロードしても問題ありません
- 既存のdata.jsデータは保持され、新しいデータで上書きされます
- ファイルサイズが大きすぎる場合は処理に時間がかかることがあります
