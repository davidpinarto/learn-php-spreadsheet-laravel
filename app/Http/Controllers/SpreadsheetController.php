<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class SpreadsheetController extends Controller
{
    public function postSpreadsheet()
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        /**
         * ingat! 
         * huruf merepresentasikan kolom
         * angka merepresentasikan baris
         * dan cell adalah pertemuan antara kolom dan baris (A1, A5)
         */
        $sheet->setCellValue('A1', 'Judul Pertama'); // kolom 1 row 1
        $sheet->setCellValue('B1', 'Judul Kedua'); // kolom 2 row 1

        $character = str_split('ABCDEFGHIJKLMNOPQRSTUVWXYZ');
        $header = [
            'NIK',
            'Nama Lengkap',
            'Nomor Handphone',
            'Alamat',
            'Tanggal Lahir'
        ];

        /**
         * bermain dengan mengisi semua data di kolom judul pertama :D
         */
        // for ( $i = 2; $i < 20; $i++ ) {
        //     $sheet->setCellValue("A$i", "isi ke $i");
        // }

        /**
         * set judul di tiap baris 1
         */
        foreach ($header as $index => $title) {
            $columnLetter = $character[$index]; 
            $cell = $columnLetter . 1;
            $sheet->setCellValue($cell, $title);
        }

        $writer = new Xlsx($spreadsheet);

        $filePath = storage_path('app/public/testing.xlsx');
        $writer->save($filePath);

        $response = [
            'status' => 'success',
            'message' => 'Create new spreadsheet success!'
        ];
        return response()->json($response, 201);
    }
}
