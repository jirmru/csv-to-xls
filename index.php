<?php

// Composerのオートローダーを読み込む
// この行が動作するためには、`composer install` を実行している必要があります。
require 'vendor/autoload.php';

// PhpSpreadsheetのクラスをインポート
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

/**
 * CSVテキストをExcel（.xlsx）ファイルとして出力する関数
 *
 * @param string $csv_text      日本語を含むCSV形式のテキストデータ
 * @param string $filename      出力するファイル名 (例: 'report.xlsx')
 * @param bool   $has_header    trueの場合、CSVの1行目を見出し行として太字で装飾する
 * @return void
 */
function csv_to_xlsx(string $csv_text, string $filename = 'export.xlsx', bool $has_header = true): void
{
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

    // Spreadsheetオブジェクトの作成
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // データをセルに書き込む
    $row_index = 1;
    foreach ($data as $row_data) {
        $col_index = 1;
        foreach ($row_data as $cell_data) {
            // 全てのセルを文字列として明示的に設定
            // これにより '001' のような値が数値の 1 になるのを防ぐ
            $sheet->setCellValueExplicitByColumnAndRow(
                $col_index,
                $row_index,
                $cell_data,
                DataType::TYPE_STRING
            );
            $col_index++;
        }
        $row_index++;
    }

    // ヘッダー行を太字にする
    if ($has_header && count($data) > 0) {
        $header_range = 'A1:' . $sheet->getHighestColumn() . '1';
        $sheet->getStyle($header_range)->getFont()->setBold(true);
    }

    // 各列の幅を自動調整
    foreach ($sheet->getColumnIterator() as $column) {
        $sheet->getColumnDimension($column->getColumnIndex())->setAutoSize(true);
    }

    // HTTPヘッダーを設定してExcelファイルとしてダウンロードさせる
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment; filename="' . rawurlencode($filename) . '"');
    header('Cache-Control: max-age=0');

    // Writerを作成し、ブラウザに出力
    $writer = new Xlsx($spreadsheet);
    $writer->save('php://output');

    exit;
}

// --- 以下、関数の使用例 ---

// このファイルがWebサーバー経由で直接アクセスされた場合
if (php_sapi_name() !== 'cli') {

    // サンプルCSVデータ (日本語、特殊文字、カンマ、改行を含む)
    // 商品ID '001' のような値を正しく扱う例を追加
    $sample_csv = <<<CSV
"製品ID","製品名","カテゴリ","価格","在庫数"
"001","高性能ノートPC","コンピュータ",150000,50
"002","ワイヤレスマウス","アクセサリ",3500,"200"
"003","4Kモニター, 27インチ","ディスプレイ",45000,30
"004","メカニカルキーボード (青軸)","アクセサリ",12000,100
CSV;

    // 関数を呼び出し
    csv_to_xlsx($sample_csv, '製品リスト.xlsx', true);
}
