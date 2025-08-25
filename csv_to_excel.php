<?php

// Composerのオートローダーを読み込む
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

/**
 * CSVテキストをExcel（.xlsx）ファイルとして出力する関数
 *
 * @param string $csv_text      日本語を含むCSV形式のテキストデータ
 * @param string $filename      出力するファイル名 (例: 'report.xlsx')
 * @param bool   $has_header    trueの場合、CSVの1行目を見出し行として装飾する
 * @return void
 */
function csv_to_xlsx(string $csv_text, string $filename = 'export.xlsx', bool $has_header = true): void
{
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
        return; // データがなければ何もしない
    }

    // Spreadsheetオブジェクトを作成
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // データをセルに書き込む
    $rowIndex = 1;
    foreach ($data as $row) {
        $colIndex = 1;
        foreach ($row as $cell) {
            $cellCoordinate = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colIndex) . $rowIndex;

            // '0'で始まる値が数値に変換されるのを防ぐ
            if (is_string($cell) && strlen($cell) > 1 && $cell[0] === '0' && !is_numeric(substr($cell, 1))) {
                 $sheet->getCell($cellCoordinate)->setValueExplicit($cell, DataType::TYPE_STRING);
            } else {
                 $sheet->getCell($cellCoordinate)->setValue($cell);
            }
            $colIndex++;
        }
        $rowIndex++;
    }

    // ヘッダー行を装飾
    if ($has_header && count($data) > 0) {
        $header_range = 'A1:' . \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex(count($data[0])) . '1';
        $sheet->getStyle($header_range)->getFont()->setBold(true);
        $sheet->getStyle($header_range)->getFill()
            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
            ->getStartColor()->setARGB('FFF0F0F0');
    }

    // 列の幅を自動調整
    foreach (range(1, count($data[0])) as $col) {
        $sheet->getColumnDimension(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col))->setAutoSize(true);
    }


    // HTTPヘッダーを設定
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    // ファイル名をURLエンコードし、日本語ファイル名に対応
    $encoded_filename = rawurlencode($filename);
    header('Content-Disposition: attachment;filename="' . $encoded_filename . '"; filename*=UTF-8\'\'' . $encoded_filename);
    header('Cache-Control: max-age=0');

    // Writerを作成して出力
    $writer = new Xlsx($spreadsheet);
    $writer->save('php://output');

    exit;
}

// --- 以下、関数の使用例 ---

// このファイルがWebサーバー経由で直接アクセスされた場合、またはCLIで実行された場合
if (php_sapi_name() === 'cli' || isset($_SERVER['REQUEST_URI'])) {

    // サンプルCSVデータ (日本語、特殊文字、カンマ、改行を含む)
    $sample_csv = <<<CSV
"製品名","カテゴリ","価格","在庫数","製品コード"
"高性能ノートPC","コンピュータ",150000,50,"PC-001"
"ワイヤレスマウス","アクセサリ",3500,"200","AC-002"
"4Kモニター, 27インチ","ディスプレイ",45000,30,"DS-003"
"メカニカルキーボード (青軸)","アクセサリ",12000,100,"AC-004"
"01番のサンプル","テスト",100,10,"01-TEST"
CSV;

    // 関数名を修正し、ファイル名を.xlsxに変更
    csv_to_xlsx($sample_csv, '製品リスト.xlsx', true);
}
