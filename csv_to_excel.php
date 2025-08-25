<?php

/**
 * CSVテキストをExcel（.xlsx）ファイルとして出力する関数
 *
 * 注意: この関数は厳密な.xlsxバイナリを生成するのではなく、
 * Excelが解釈可能なHTMLテーブルを生成します。
 * これにより、外部ライブラリなしで動作し、日本語の文字化けを防ぎます。
 *
 * @param string $csv_text      日本語を含むCSV形式のテキストデータ
 * @param string $filename      出力するファイル名 (例: 'report.xlsx')
 * @param bool   $has_header    trueの場合、CSVの1行目を見出し行として太字で装飾する
 * @return void
 */
function csv_to_xlsx(string $csv_text, string $filename = 'export.xlsx', bool $has_header = true): void
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

    // HTTPヘッダーを設定してExcelファイルとしてダウンロードさせる
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    $user_agent = $_SERVER['HTTP_USER_AGENT'] ?? '';
    $encoded_filename = rawurlencode($filename);
    if (preg_match('/(MSIE|Trident)/', $user_agent)) {
        // IE(Trident)用の処理
        header('Content-Disposition: attachment; filename="' . $encoded_filename . '"');
    } else {
        // モダンブラウザ用の処理 (RFC 6266)
        header('Content-Disposition: attachment; filename*=UTF-8\'\'' . $encoded_filename);
    }
    header('Cache-Control: max-age=0');

    // Excelでの文字化けを確実に防ぐため、BOMを先頭に付与
    echo "\xEF\xBB\xBF";

    // HTMLテーブルとして出力
    echo '<html><head><meta charset="UTF-8"></head><body>';
    echo '<table>';

    $is_first_row = true;
    foreach ($data as $row) {
        echo '<tr>';
        if ($has_header && $is_first_row) {
            // 見出し行の処理
            foreach ($row as $cell) {
                echo '<th style="font-weight: bold; background-color: #f0f0f0;">' . htmlspecialchars($cell, ENT_QUOTES, 'UTF-8') . '</th>';
            }
            $is_first_row = false;
        } else {
            // データ行の処理
            foreach ($row as $cell) {
                // '0'で始まる商品コードなどが数値に変換されるのを防ぐため、styleでmso-number-formatを指定
                $style = is_numeric($cell) && strlen($cell) < 15 ? '' : 'style="mso-number-format:\'@\'"';
                echo '<td ' . $style . '>' . htmlspecialchars($cell, ENT_QUOTES, 'UTF-8') . '</td>';
            }
        }
        echo '</tr>';
    }

    echo '</table>';
    echo '</body></html>';

    // 処理を終了
    exit;
}

// --- 以下、関数の使用例 ---

// このファイルがWebサーバー経由で直接アクセスされた場合のみ、
// 以下のサンプルコードが実行され、Excelファイルがダウンロードされます。
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
    // 第3引数: 1行目を見出しとして扱うか (true: はい, false: いいえ)
    csv_to_xlsx($sample_csv, '製品リスト.xlsx', true);
}
