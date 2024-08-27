<?php

namespace App\Http\Controllers;

use Carbon\Carbon;
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

        $employeesData = [
            [
                'NIK' => 1,
                'fullname' => 'David Pinarto',
                'phone_number' => '085394829835',
                'address' => 'Sulawesi selatan',
                'birth_of_date' => Carbon::now()->format('Y-m-d'),
            ],
            [
                'NIK' => 2,
                'fullname' => 'Eko Sunandar',
                'phone_number' => '0859237423',
                'address' => 'Lampung',
                'birth_of_date' => Carbon::now()->format('Y-m-d'),
            ],
            [
                'NIK' => 3,
                'fullname' => 'Erik Juniadi',
                'phone_number' => '085973782324',
                'address' => 'Semarang',
                'birth_of_date' => Carbon::now()->format('Y-m-d'),
            ]
        ];

        /** rowNumber adalah nomor baris */
        $rowNumber = 2;
        /** index ini berfungsi untuk mengambil value dari character dimulai dari 0 [A, B ,C] */
        $index = 0;
        /** Lakukan perulangan pada nested array employeesData */
        foreach ($employeesData as $_ => $employeeData) {
            
            /** Lakukan perulangan untuk setiap value nilai nested array */
            foreach ($employeeData as $_ => $employeeData2) {
                $columnLetter = $character[$index];
                $cell = $columnLetter . $rowNumber;
                $sheet->setCellValue($cell, $employeeData2);
                
                /** mengisi nilai cell kolom A telah selesai, lanjutt isi kolom B */
                $index += 1;
            }

            /** set index kembali ke 0 yaitu huruf A, karena perulangan pada value employee data pertama telah selesai */
            $index = 0;
            /** kita menuju ke baris berikutnya untuk mengisi nilai employee data kedua */
            $rowNumber += 1;
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
