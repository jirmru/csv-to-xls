<?php

/**
 * CSVテキストをExcel（.xls）ファイルとして出力する関数
 *
 * XML Spreadsheet 2003 形式で出力することで、
 * 拡張子(.xls)と内容の不一致によるExcelの警告を解消します。
 * この方法は外部ライブラリを必要としません。
 *
 * @param string $csv_text      日本語を含むCSV形式のテキストデータ
 * @param string $filename      出力するファイル名 (例: 'report.xls')
 * @param bool   $has_header    trueの場合、CSVの1行目を見出し行として太字で装飾する
 * @return void
 */
function csv_to_xls(string $csv_text, string $filename = 'export.xls', bool $has_header = true): void
{
    // 内部処理をUTF-8に統一
    mb_internal_encoding('UTF-8');

    // CSVデータを解析
    $csv_text = str_replace(["\r\n", "\r"], "\n", trim($csv_text));
    $lines = explode("\n", $csv_text);
    $data = [];
    foreach ($lines as $line) {
        if (!empty($line)) {
            $data[] = str_getcsv($line);
        }
    }

    if (empty($data)) {
        return;
    }

    // HTTPヘッダーを設定
    header('Content-Type: application/vnd.ms-excel; charset=UTF-8');
    header('Content-Disposition: attachment; filename="' . rawurlencode($filename) . '"');
    header('Cache-Control: max-age=0');

    // XML Spreadsheet 2003 形式のヘッダーを出力
    echo '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
    echo '<?mso-application progid="Excel.Sheet"?>' . "\n";
    echo '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"' . "\n";
    echo ' xmlns:o="urn:schemas-microsoft-com:office:office"' . "\n";
    echo ' xmlns:x="urn:schemas-microsoft-com:office:excel"' . "\n";
    echo ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"' . "\n";
    echo ' xmlns:html="http://www.w3.org/TR/REC-html40">' . "\n";

    // スタイルの定義
    echo ' <Styles>' . "\n";
    // 通常スタイル
    echo '  <Style ss:ID="Default" ss:Name="Normal">' . "\n";
    echo '   <Alignment ss:Vertical="Bottom"/>' . "\n";
    echo '   <Borders/>' . "\n";
    echo '   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>' . "\n";
    echo '   <Interior/>' . "\n";
    echo '   <NumberFormat/>' . "\n";
    echo '   <Protection/>' . "\n";
    echo '  </Style>' . "\n";
    // ヘッダースタイル (太字)
    echo '  <Style ss:ID="sHeader">' . "\n";
    echo '   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000" ss:Bold="1"/>' . "\n";
    echo '  </Style>' . "\n";
    // 文字列スタイル (数値の自動変換を防ぐ)
    echo '  <Style ss:ID="sText">' . "\n";
    echo '   <NumberFormat ss:Format="@"/>' . "\n";
    echo '  </Style>' . "\n";
    echo ' </Styles>' . "\n";

    // ワークシートの開始
    echo ' <Worksheet ss:Name="Sheet1">' . "\n";
    echo '  <Table>' . "\n";

    $is_first_row = true;
    foreach ($data as $row) {
        echo '   <Row>' . "\n";
        if ($has_header && $is_first_row) {
            // 見出し行
            foreach ($row as $cell) {
                echo '    <Cell ss:StyleID="sHeader"><Data ss:Type="String">' . htmlspecialchars($cell, ENT_QUOTES, 'UTF-8') . '</Data></Cell>' . "\n";
            }
            $is_first_row = false;
        } else {
            // データ行
            foreach ($row as $cell) {
                // 15桁以上の数字はExcelで精度が落ちるため、文字列として扱う
                if (is_numeric($cell) && strlen($cell) < 15) {
                    echo '    <Cell><Data ss:Type="Number">' . htmlspecialchars($cell, ENT_QUOTES, 'UTF-8') . '</Data></Cell>' . "\n";
                } else {
                    echo '    <Cell ss:StyleID="sText"><Data ss:Type="String">' . htmlspecialchars($cell, ENT_QUOTES, 'UTF-8') . '</Data></Cell>' . "\n";
                }
            }
        }
        echo '   </Row>' . "\n";
    }

    // ワークシートとワークブックの終了
    echo '  </Table>' . "\n";
    echo ' </Worksheet>' . "\n";
    echo '</Workbook>' . "\n";

    // 処理を終了
    exit;
}


// --- 以下、関数の使用例 ---

// このファイルがWebサーバー経由で直接アクセスされた場合のみ、
// 以下のサンプルコードが実行され、Excelファイルがダウンロードされます。
if (isset($_SERVER['REQUEST_URI'])) {

    // サンプルCSVデータ (日本語、特殊文字、カンマ、改行を含む)
    $sample_csv = <<<CSV
"製品名","カテゴリ","価格","在庫数","商品コード"
"高性能ノートPC","コンピュータ",150000,50,"PC-001"
"ワイヤレスマウス","アクセサリ",3500,"200","AC-055"
"4Kモニター, 27インチ","ディスプレイ",45000,30,"DP-300"
"メカニカルキーボード (青軸)","アクセサリ",12000,100,"KB-112"
"00123","テスト",100,10,"00123"
CSV;

    // 関数を呼び出し
    csv_to_xls($sample_csv, '製品リスト.xls', true);
}
