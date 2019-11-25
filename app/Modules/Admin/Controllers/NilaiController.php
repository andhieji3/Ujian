<?php

namespace App\Modules\Admin\Controllers;

use Illuminate\Http\Request;
use App\Http\Controllers\Controller;
use App\Models\Sekolah;
use App\Models\Tes;
use App\Models\JadwalUjian;
use App\Models\Soal;
use App\Models\ItemSoal;
use App\Models\Siswa;
use App\Modules\Admin\Helper;

use GuzzleHttp\Client;
use Spreadsheet;
use Xlsx;
use PDF;

class NilaiController extends Controller
{

    public function index()
    {
      $jadwal = JadwalUjian::with('tes')->paginate(10)->appends(request()->except('page'));
      return view("Admin::nilai.index",[
        'title'=>'Nilai Ujian - Administrator',
        'breadcrumb'=>'Nilai Ujian',
        'jadwal'=>$jadwal,
        'tes'=>new Tes
      ]);
    }

    public function createColumnsArray($end_column, $first_letters = '')
    {
      $columns = array();
      $length = strlen($end_column);
      $letters = range('A', 'Z');

      // Iterate over 26 letters.
      foreach ($letters as $letter) {
          // Paste the $first_letters before the next.
          $column = $first_letters . $letter;

          // Add the column to the final array.
          $columns[] = $column;

          // If it was the end column that was added, return the columns.
          if ($column == $end_column)
              return $columns;
      }

      // Add the column children.
      foreach ($columns as $column) {
          // Don't itterate if the $end_column was already set in a previous itteration.
          // Stop iterating if you've reached the maximum character length.
          if (!in_array($end_column, $columns) && strlen($column) < $length) {
              $new_columns = $this->createColumnsArray($end_column, $column);
              // Merge the new columns which were created with the final columns array.
              $columns = array_merge($columns, $new_columns);
          }
      }

      return $columns;
    }

    public function downloadExcel($uuid)
    {
      $jadwal = JadwalUjian::where('uuid',$uuid)->first();

      $filename = 'Nilai '.$jadwal->nama_ujian.'.xlsx';

      $spreadsheet = new Spreadsheet();
  		$sheet = $spreadsheet->getActiveSheet();

  		$sheet->setCellValue('A1', 'No');
  		$sheet->setCellValue('B1', 'No. Ujian');
  		$sheet->setCellValue('C1', 'Nama Siswa');
      $sheet->getStyle('A1:C1')->getFont()->setBold(true);
      $sheet->getStyle('A1:D1')->getAlignment()->setHorizontal('center');
      $sheet->getStyle('A1:D1')->getAlignment()->setVertical('center');
      $sheet->getStyle('A')->getAlignment()->setHorizontal('center');

      $sheet->getColumnDimension('A')->setWidth(5);
      $sheet->getColumnDimension('B')->setWidth(20);
      $sheet->getColumnDimension('C')->setWidth(30);

      $cols = $this->createColumnsArray('ZZ');
      $cols = array_slice($cols,3,count($cols));

      // if ($jadwal->jenis_soal == 'E') {
      //   return $this->downloadEssay($uuid);
      // }

      $sheet->setCellValue('D1', 'Jumlah Soal');
      $sheet->getStyle('D1')->getFont()->setBold(true);
      $sheet->getStyle('D1')->getAlignment()->setHorizontal('center');
      $sheet->getStyle('D1')->getAlignment()->setVertical('center');
      $sheet->getColumnDimension('D')->setAutoSize(true);

      if ($jadwal->jenis_soal=='P') {
        $sheet->setCellValue('E1', 'Benar');
        $sheet->getStyle('E1')->getFont()->setBold(true);
        $sheet->getStyle('E1')->getAlignment()->setHorizontal('center');
        $sheet->getStyle('E1')->getAlignment()->setVertical('center');
        $sheet->getColumnDimension('E')->setAutoSize(true);

        $sheet->setCellValue('F1', 'Salah');
        $sheet->getStyle('F1')->getFont()->setBold(true);
        $sheet->getStyle('F1')->getAlignment()->setHorizontal('center');
        $sheet->getStyle('F1')->getAlignment()->setVertical('center');
        $sheet->getColumnDimension('F')->setAutoSize(true);
      }

      $sheet->setCellValue('G1', 'Nilai Akhir');
      $sheet->getStyle('G1')->getFont()->setBold(true);
      $sheet->getStyle('G1')->getAlignment()->setHorizontal('center');
      $sheet->getStyle('G1')->getAlignment()->setVertical('center');
      $sheet->getColumnDimension('G')->setAutoSize(true);

      $peserta = Siswa::whereIn('uuid',json_decode($jadwal->peserta))
      ->orderBy('id','asc')
      ->get();
      $i=1;
      foreach ($peserta as $key => $v) {
        $sheet->setCellValue('A'.($i+1), $i);
        $sheet->setCellValue('B'.($i+1), $v->noujian);
        $sheet->setCellValue('C'.($i+1), $v->nama);
        $benar = 0;
        $nilai = 0;

        $login = $v->attemptLogin()->where('pin',$jadwal->pin)->first();
        if ($login && $login->soal_ujian != '' && !is_null($login->soal_ujian)) {
          $soal = ItemSoal::whereIn('uuid',json_decode($login->soal_ujian))->get();
          foreach ($soal as $key1 => $s) {
            $tes = Tes::where('noujian',$v->noujian)->where('soal_item',$s->uuid)->where('pin',$jadwal->pin)->first();

            if ($tes && array_key_exists($tes->jawaban,json_decode($s->opsi))) {
              if ((string) $tes->jawaban == (string) $s->benar) {
                $benar++;
              }
            }

          }

          if ($soal->count()) {
            $nilai = 0;
          }

          if ($benar) {
            $nilai += round($benar/$jadwal->jumlah_soal*$jadwal->bobot,2);
          }

        }

        $sheet->getStyle('D'.($i+1))->getAlignment()->setHorizontal('center');
        $sheet->setCellValue('D'.($i+1),$jadwal->jumlah_soal);

        if ($jadwal->jenis_soal=='P') {
          $sheet->getStyle('E'.($i+1))->getAlignment()->setHorizontal('center');
          $sheet->setCellValue('E'.($i+1),$benar);

          $sheet->getStyle('F'.($i+1))->getAlignment()->setHorizontal('center');
          $sheet->setCellValue('F'.($i+1),$jadwal->jumlah_soal-$benar);
        }

        $sheet->getStyle('G'.($i+1))->getAlignment()->setHorizontal('center');
        $sheet->getStyle('G'.($i+1))->getFont()->setBold(true);
        $sheet->setCellValue('G'.($i+1),$nilai);

        $i++;
      }

      $writer = new Xlsx($spreadsheet);
  		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      header('Content-Disposition: attachment; filename="'.$filename.'"');
  		$writer->save("php://output");
    }

    public function downloadPDF($uuid)
    {
      $jadwal = JadwalUjian::where('uuid',$uuid)->first();

      $filename = 'Nilai '.$jadwal->nama_ujian.'.pdf';

      $peserta = Siswa::whereIn('uuid',json_decode($jadwal->peserta))
      ->orderBy('id','asc')
      ->get();

      $pdf = PDF::loadView('Admin::nilai.nilai-pdf',[
        'jadwal'=>$jadwal,
        'peserta'=>$peserta,
        'title'=>'Nilai '.$jadwal->nama_ujian,
        'sekolah'=>Sekolah::first(),
        'helper'=>new Helper
      ]);

      return $pdf->setOption('page-width','21.5cm')
      ->setOption('page-height','33cm')
      // ->setOption('margin-bottom',0)
      // ->setOption('margin-left',0)
      // ->setOption('margin-right',0)
      ->stream($filename);

    }

    public function detail($uuid)
    {
      $jadwal = JadwalUjian::with('tes')->where('uuid',$uuid)->first();
      return view("Admin::nilai.detail",[
        'title'=>'Nilai '.$jadwal->nama_ujian.' - Administrator',
        'breadcrumb'=>'Nilai Ujian',
        'jadwal'=>$jadwal,
        'peserta'=>Siswa::whereIn('uuid',json_decode($jadwal->peserta))->orderBy('id','asc')->get(),
      ]);
    }

    public function detailDownload($ujian,$siswa)
    {
      $nilai = 0;
      $nbenar = 0;
      $jadwal = JadwalUjian::with('tes')->where('uuid',$ujian)->first();
      $siswa = Siswa::where('uuid',$siswa)->first();
      $plogin = $siswa->attemptLogin()->where('pin',$jadwal->pin)->first();
      $jumlah_soal = count(json_decode($plogin->soal_ujian));
      $dtes = Tes::where('noujian',$siswa->noujian)
      ->where('pin',$jadwal->pin)->whereIn('soal_item',json_decode($plogin->soal_ujian))->get();

      $soal = [];
      $siswaSoal = json_decode($plogin->soal_ujian);
      if (count($siswaSoal)) {
        foreach ($siswaSoal as $key => $s) {
          $gs = ItemSoal::where('uuid',$s)->first();
          if ($gs) {
            array_push($soal,$gs);
          }
        }
      }
      foreach ($dtes as $key => $tes) {
        $benar = $tes->soalItem->benar;
        if (!is_null($benar) && (string) $tes->jawaban == (string) $benar && $tes->soalItem->jenis_soal=='P') {
          $nbenar++;
        }
      }
      if ($jumlah_soal) {
        $nilai = 0;
      }
      if ($nbenar) {
        $nilai += round($nbenar/$jumlah_soal*$jadwal->bobot,2);
      }
      // return view("Admin::nilai.detail-download",[
      //   'title'=>'Nilai '.$jadwal->nama_ujian.' - ('.$siswa->noujian.') '.$siswa->nama,
      //   'breadcrumb'=>'Nilai Ujian',
      //   'jadwal'=>$jadwal,
      //   'siswa'=>$siswa,
      //   'soal'=>$soal,
      //   'nilai'=>$nilai,
      //   'sekolah'=>Sekolah::first(),
      //   'helper'=>new Helper
      // ]);

      $filename = 'Nilai '.$jadwal->nama_ujian.'.pdf';

      $pdf = PDF::loadView("Admin::nilai.detail-download",[
        'title'=>'Nilai '.$jadwal->nama_ujian.' - ('.$siswa->noujian.') '.$siswa->nama,
        'breadcrumb'=>'Nilai Ujian',
        'jadwal'=>$jadwal,
        'siswa'=>$siswa,
        'soal'=>$soal,
        'nilai'=>$nilai,
        'nilai'=>$nilai,
        'benar'=>$nbenar,
        'sekolah'=>Sekolah::first(),
        'helper'=>new Helper
      ]);

      return $pdf->setOption('page-width','21.5cm')
      ->setOption('page-height','33cm')
      // ->setOption('margin-bottom',0)
      // ->setOption('margin-left',0)
      // ->setOption('margin-right',0)
      ->stream($filename);

    }

}
