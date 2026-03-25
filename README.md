# kasouhin_uriagezumi_report（架装品実績レポート・売上）

## 依頼・リポジトリ

- **GitHub:** [https://github.com/dezigozi/kasouhin_uriagezumi_report.git](https://github.com/dezigozi/kasouhin_uriagezumi_report.git)
- **プッシュ対象（ローカル）:**  
  `E:\Dropbox\会社\仕事\マイドキュメント\Ｍ／G\生産性向上(働き方改革)\AI\●仕事\架装品実績レポート\売上`

上記リポジトリに、上記 **売上** フォルダのデータをプッシュする作業を実施しました。

## 作業が完了しました

### 結果

[dezigozi/kasouhin_uriagezumi_report](https://github.com/dezigozi/kasouhin_uriagezumi_report) の **`main`** ブランチへ、指定の「売上」フォルダの内容をプッシュ済みです。

### 実施したこと

- **売上** フォルダで Git を初期化し、リモート `origin` を上記 URL に設定しました。
- **`.gitignore`** … `node_modules/` のみ除外（`npm install` で再取得できます）。依存の数万ファイルはリポジトリに含めていません。
- **GitHub の 100MB 制限対策** … 次の大きいファイルは **Git LFS** で管理しています（[Git Large File Storage](https://git-lfs.github.com/)）。
  - `*.asar`（約 137MB ×2）
  - `**/win-unpacked/*.exe`（約 169MB ×2、内容が同じため LFS オブジェクトは 3 個・合計約 464MB でアップロード）
- **`売上データ※バックアップ/粗利・売上_詳細_架装品_202312-202601.xlsx`（約 77.6MB）** … 100MB 未満のため通常の Git でコミットしています。GitHub から「50MB 推奨」を超える旨の **警告** は出ますが、プッシュは成功しています。

### クローンするとき

別 PC でクローンする場合は、**Git LFS 入りの Git** で `git clone` してください（LFS ファイルも取得されます）。

```bash
git clone https://github.com/dezigozi/kasouhin_uriagezumi_report.git
```

`node_modules` は含まれていないので、アプリを開発する場合は `lease-report-app` 内で `npm install` が必要です。

### 補足

初回プッシュ時は PowerShell で `HEAD~1` がホームパスに解釈されるため、履歴の巻き戻しに失敗し、`.git` を一度削除して LFS 付きでやり直しています。現在のリポジトリは整合しています。
