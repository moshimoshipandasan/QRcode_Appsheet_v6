# QRコード生成 Google Apps Script

## 概要

このGoogle Apps Scriptは、Googleスプレッドシートの特定のシート（デフォルトでは "t氏名"）にあるデータからQRコードを生成し、指定された列に画像として表示します。A列にデータがない場合は、ランダムな文字列を生成してQRコードを作成します。

また、別のシート（"tQR読込み"）からデータを集計し、"集計"シートに出力する機能も含まれています。

## 機能

*   **QRコード作成**:
    *   "t氏名"シートのA列のデータを元にQRコードを生成します。
    *   A列が空の場合は、ランダムな8文字の英数字を生成してQRコードを作成します。
    *   生成されたQRコードは、D列に `=IMAGE()` 関数を用いた画像として表示されます。
    *   E列は `FALSE` に設定されます。
*   **集計準備**:
    *   "tQR読込み"シートの特定の範囲（K列からO列）のデータを"集計"シートにコピーします。
    *   コピー前に"集計"シートの内容はクリアされます。

## 使い方

1.  **スクリプトの導入**:
    *   Googleスプレッドシートを開き、「ツール」>「スクリプトエディタ」を選択します。
    *   エディタに `コード.js` の内容をコピー＆ペーストします。
    *   プロジェクトに名前を付けて保存します。
2.  **シートの準備**:
    *   スプレッドシートに "t氏名", "tQR読込み", "集計" という名前のシートを作成します。
    *   "t氏名"シートのA列にQRコードにしたいデータを入力します。
3.  **初回実行と承認**:
    *   スプレッドシートを再読み込みするか、再度開くと、「QRコード」というカスタムメニューが表示されます。
    *   メニューから「QRコード作成」または「集計準備」を初めて実行する際に、スクリプトの実行に必要な承認を求められます。内容を確認し、承認してください。
4.  **機能の実行**:
    *   **QRコード作成**: 「QRコード」メニューから「QRコード作成」を選択すると、"t氏名"シートのD列にQRコードが表示されます。
    *   **集計準備**: 「QRコード」メニューから「集計準備」を選択すると、"tQR読込み"シートから"集計"シートへデータがコピーされます。

## 関連リソース

*   **スプレッドシート テンプレート**: [こちらをクリックしてコピーを作成](https://docs.google.com/spreadsheets/d/1NpsZOMLb7FnpNNZBbctQn4IyTRYV7zLb-lWJx5OvtGQ/copy)
*   **解説動画 (YouTube)**: [視聴する](https://youtu.be/LZ3yVW1OHfk?si=L5KllhINA9Yfyaid)

## 注意事項

*   QRコードの生成には外部API ([QR Server API](https://goqr.me/api/)) を使用しています。APIの仕様変更やサービス停止により、機能が利用できなくなる可能性があります。
*   スクリプトの実行にはGoogleアカウントが必要です。
*   大量のデータを処理する場合、Google Apps Scriptの実行時間制限に達する可能性があります。

## ライセンス

このスクリプトは [MIT License](LICENSE) の下で公開されています。

## 作成者

Noboru Ando @ Aoyama Gakuin University
(元コード作成者: 青山学院中等部 安藤昇)
