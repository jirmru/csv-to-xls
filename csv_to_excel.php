<?php

/**
 * CSVテキストを標準的なCSVファイルとして出力する関数
 *
 * この関数は、入力されたCSV文字列を解析し、
 * UTF-8エンコードされたCSVファイルとしてブラウザにダウンロードさせます。
 * - 日本語のファイル名が文字化けしないようにContent-Dispositionヘッダーを調整します。
 * - Excelでファイルを開いた際に文字化けを防ぐため、BOM（Byte Order Mark）を付与します。
 * - fputcsvを使用して、RFC 4180に準拠したCSVを生成します。
 *
 * @param string $csv_text      日本語を含むCSV形式のテキストデータ
 * @param string $filename      出力するファイル名 (例: 'report.csv')
 * @return void
 */
function generate_csv_output(string $csv_text, string $filename = 'export.csv'): void
{
    // 内部処理をUTF-8に統一
    mb_internal_encoding('UTF-8');

    // CSVデータを解析
    // 改行コードの揺れを吸収 (CRLF, LF, CR -> LF)
    $csv_text = str_replace(["\r\n", "\r"], "\n", trim($csv_text));
    $lines = explode("\n", $csv_text);
    $data = [];
    foreach ($lines as $line) {
        if (!empty($line)) {
            // ダブルクォートで囲まれたカンマも正しく扱えるようにstr_getcsvを使用
            $data[] = str_getcsv($line);
        }
    }

    if (empty($data)) {
        return; // データがなければ何もしない
    }

    // HTTPヘッダーを設定してCSVファイルとしてダウンロードさせる
    header('Content-Type: text/csv; charset=UTF-8');

    // ファイル名のエンコーディング問題を解決するため、RFC 6266に準拠したヘッダーを使用
    $encoded_filename = rawurlencode($filename);
    // モダンブラウザ向けのfilename*と、古いブラウザ向けのfilenameを提供
    header('Content-Disposition: attachment; filename*=UTF-8\'\'' . $encoded_filename . '; filename="' . $filename . '"');

    header('Cache-Control: max-age=0');

    // メモリストリームを開いてCSVデータを書き込む
    $handle = fopen('php://memory', 'w');
    if ($handle === false) {
        // @codeCoverageIgnoreStart
        // In a normal environment, this should not fail.
        // But as a good practice, we handle the error case.
        header("HTTP/1.1 500 Internal Server Error");
        echo "Failed to create memory stream.";
        exit;
        // @codeCoverageIgnoreEnd
    }

    // Excelでの文字化けを確実に防ぐため、BOMを先頭に付与
    fwrite($handle, "\xEF\xBB\xBF");

    // fputcsvでデータを書き込む
    foreach ($data as $row) {
        fputcsv($handle, $row);
    }

    // ストリームの先頭にポインタを戻す
    rewind($handle);

    // ストリームの内容を読み込んで出力
    echo stream_get_contents($handle);

    // ハンドルを閉じる
    fclose($handle);

    // 処理を終了
    exit;
}

// --- 以下、関数の使用例 ---

// このファイルがWebサーバー経由で直接アクセスされた場合のみ、
// 以下のサンプルコードが実行され、CSVファイルがダウンロードされます。
if (isset($_SERVER['REQUEST_URI'])) {

    // サンプルCSVデータ (日本語、特殊文字、カンマ、改行を含む)
    $sample_csv = <<<CSV
"製品名","カテゴリ","価格","在庫数"
"高性能ノートPC","コンピュータ",150000,50
"ワイヤレスマウス","アクセサリ",3500,"200"
"4Kモニター, 27インチ","ディスプレイ",45000,30
"メカニカルキーボード (青軸)","アクセサリ",12000,100
CSV;

    // 関数を呼び出し
    // 第1引数: CSVテキスト
    // 第2引数: 出力ファイル名
    generate_csv_output($sample_csv, '製品リスト.csv');
}
