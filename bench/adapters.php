<?php

declare(strict_types=1);

/*
 * One adapter per library, each writing/reading the identical workload so
 * the comparison is apples-to-apples. Every adapter uses the library's own
 * recommended fast path (streaming APIs where they exist).
 */

return [
    'easy-excel' => [
        'write' => function (string $file, int $rows): void {
            \EasyExcel\Native::assertAvailable();
            $s = new \EasyExcel\Compat\Spreadsheet();
            $ws = $s->getActiveSheet();
            $chunk = [];
            $at = 1;
            foreach (benchRows($rows) as $row) {
                $chunk[] = $row;
                if (\count($chunk) === 2048) {
                    $ws->fromArray($chunk, null, 'A' . $at, true);
                    $at += 2048;
                    $chunk = [];
                }
            }
            if ($chunk) {
                $ws->fromArray($chunk, null, 'A' . $at, true);
            }
            (new \EasyExcel\Compat\Writer\Xlsx($s))->save($file);
            $s->disconnectWorksheets();
        },
        'write-styled' => function (string $file, int $rows): void {
            \EasyExcel\Native::assertAvailable();
            benchStyledWrite(new \EasyExcel\Compat\Spreadsheet(), \EasyExcel\Compat\Writer\Xlsx::class, $file, $rows);
        },
        'read' => function (string $file): int {
            \EasyExcel\Native::assertAvailable();
            $s = (new \EasyExcel\Compat\Reader\Xlsx())->load($file);
            $cells = 0;
            foreach ($s->getActiveSheet()->toArray() as $row) {
                $cells += \count($row);
            }
            $s->disconnectWorksheets();

            return $cells;
        },
    ],

    'phpspreadsheet' => [
        'write' => function (string $file, int $rows): void {
            $s = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
            $ws = $s->getActiveSheet();
            $r = 1;
            foreach (benchRows($rows) as $row) {
                $ws->fromArray($row, null, 'A' . $r++, true);
            }
            (new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($s))->save($file);
            $s->disconnectWorksheets();
        },
        'write-styled' => function (string $file, int $rows): void {
            benchStyledWrite(new \PhpOffice\PhpSpreadsheet\Spreadsheet(), \PhpOffice\PhpSpreadsheet\Writer\Xlsx::class, $file, $rows);
        },
        'read' => function (string $file): int {
            $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
            $reader->setReadDataOnly(true);
            $s = $reader->load($file);
            $cells = 0;
            foreach ($s->getActiveSheet()->toArray() as $row) {
                $cells += \count($row);
            }
            $s->disconnectWorksheets();

            return $cells;
        },
    ],

    'openspout' => [
        'write' => function (string $file, int $rows): void {
            $writer = new \OpenSpout\Writer\XLSX\Writer();
            $writer->openToFile($file);
            foreach (benchRows($rows) as $row) {
                $writer->addRow(\OpenSpout\Common\Entity\Row::fromValues($row));
            }
            $writer->close();
        },
        'read' => function (string $file): int {
            $reader = new \OpenSpout\Reader\XLSX\Reader();
            $reader->open($file);
            $cells = 0;
            foreach ($reader->getSheetIterator() as $sheet) {
                foreach ($sheet->getRowIterator() as $row) {
                    $cells += \count($row->getCells());
                }
            }
            $reader->close();

            return $cells;
        },
    ],

    'fast-excel-writer' => [
        'write' => function (string $file, int $rows): void {
            $excel = \avadim\FastExcelWriter\Excel::create(['Sheet1']);
            $sheet = $excel->sheet();
            foreach (benchRows($rows) as $row) {
                $sheet->writeRow($row);
            }
            $excel->save($file);
        },
    ],

    'rap2hpoutre' => [
        'write' => function (string $file, int $rows): void {
            $gen = static function () use ($rows): \Generator {
                foreach (benchRows($rows) as $row) {
                    yield $row;
                }
            };
            (new \Rap2hpoutre\FastExcel\FastExcel($gen()))->export($file);
        },
        'read' => function (string $file): int {
            $cells = 0;
            (new \Rap2hpoutre\FastExcel\FastExcel())->import($file, function (array $row) use (&$cells): void {
                $cells += \count($row);
            });

            return $cells;
        },
    ],
];
