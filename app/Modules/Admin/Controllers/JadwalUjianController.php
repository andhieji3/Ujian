<?php

namespace App\Modules\Admin\Controllers;

use Illuminate\Http\Request;
use App\Http\Controllers\Controller;
use App\Models\Soal;
use App\Models\JadwalUjian;
use App\Models\Tes;
use App\Models\Login;
use App\Models\Sekolah;
use App\Models\Siswa;
use App\Models\Kelas;
use Illuminate\Support\Str;
use Carbon\Carbon;
use App\Modules\Admin\Helper;
use Auth;

use GuzzleHttp\Client;
use Validator;
use PDF;

class JadwalUjianController extends Controller
{

    public function index(Request $r)
    {
      $jadwalUjian = JadwalUjian::when($r->cari,function($q,$role){
        $role = '%'.$role.'%';
        $q->where('kode','ilike',$role)
        ->orWhere('nama','ilike',$role)
        ->orWhereHas('getSoal',function($q) use($role){
          $q->where('nama','ilike',$role);
        });
      })
      ->orderBy('aktif','desc')
      ->orderBy('updated_at','desc')
      ->paginate(10)->appends(request()->except('page'));
      $data = [
        'title' => 'Jadwal Ujian - Administrator',
        'breadcrumb' => 'Jadwal Ujian',
        'jadwal' => $jadwalUjian
      ];
      return view("Admin::master.jadwalujian.index",$data);
    }

    public function create(Request $r)
    {
      if ($r->ajax()) {
        $soal = Soal::all();
        return view("Admin::master.jadwalujian.create",[
          'soal'=>$soal,
          'kelas'=>Kelas::all()
        ]);
      }
      return redirect()->route('admin.index');
    }

    public function store(Request $r)
    {
      if ($r->ajax()) {
        $soal = Soal::where('kode',$r->kode_soal)->first();
        $valid = Validator::make($r->all(),[
          'kode_soal' => 'required',
          'kode_kelas' => 'required',
          'mulai_ujian' => 'required',
          'selesai_ujian' => 'required',
          'lama_ujian' => 'required|numeric|min:1',
          'sesi_ujian' => 'required|numeric|min:1',
          'pin_digit' => 'required|numeric|min:1',
        ],[
          'kode_soal.required' => 'Soal ujian tidak boleh kosong',
          'kode_kelas.required' => 'Kelas tidak boleh kosong',
          'mulai_ujian.required' => 'Waktu mulai ujian tidak boleh kosong',
          'mulai_ujian.date_format' => 'Format waktu mulai ujian tidak benar',
          'selesai_ujian.required' => 'Waktu selesai ujian tidak boleh kosong',
          'selesai_ujian.date_format' => 'Format waktu selesai ujian tidak benar',
          'lama_ujian.required' => 'Lama ujian tidak boleh kosong',
          'lama_ujian.numeric' => 'Lama ujian harus berupa angka dalam menit',
          'lama_ujian.min' => 'Lama ujian tidak boleh kurang dari 1 menit',
          'sesi_ujian.required' => 'Sesi ujian tidak boleh kosong',
          'sesi_ujian.numeric' => 'Sesi ujian harus berupa angka',
          'sesi_ujian.min' => 'Sesi ujian tidak boleh kurang dari 1',
          'pin_digit.required' => 'Jumlah digit pin tidak boleh kosong',
          'pin_digit.numeric' => 'Jumlah digit pin harus berupa angka',
          'pin_digit.min' => 'Jumlah digit pin tidak boleh kurang dari 1',
          ]);

          if ($valid->fails()) {
            $errs = $valid->errors()->all();
          }else{
            $errs = [];
            if ($r->mulai_ujian!=''&&$r->selesai_ujian!=''&&$r->lama_ujian!='') {
              $now = Carbon::parse(date('Y-m-d H:i'));
              $mulai_ujian = Carbon::createFromFormat('d/m/Y - H:i',$r->mulai_ujian);
              $selesai_ujian = Carbon::createFromFormat('d/m/Y - H:i',$r->selesai_ujian);
              if ($selesai_ujian < $now) {
                array_push($errs,'Waktu selesai ujian tidak boleh kurang dari waktu sekarang');
              }elseif ($mulai_ujian > $selesai_ujian) {
                array_push($errs,'Waktu selesai ujian tidak boleh kurang dari waktu mulai ujian');
              }elseif ($mulai_ujian->addMinutes($r->lama_ujian) > $selesai_ujian) {
                array_push($errs,'Lama ujian melebihi rentang waktu ujian');
              }
            }
          }
          if (count($errs)) {
            return response()->json([
              'success'=>false,
              'messages'=>$errs
            ]);
          }
          return response()->json([
            'success'=>true
          ]);
      }

      $mulai_ujian = Carbon::createFromFormat('d/m/Y - H:i',$r->mulai_ujian);
      $selesai_ujian = Carbon::createFromFormat('d/m/Y - H:i',$r->selesai_ujian);

      $jadwalUjian = new JadwalUjian;
      $jadwalUjian->uuid = (string) Str::uuid();
      $jadwalUjian->kode_soal = $r->kode_soal;
      $jadwalUjian->kode_kelas = $r->kode_kelas;
      $jadwalUjian->mulai_ujian = $mulai_ujian->toDateTimeString();
      $jadwalUjian->selesai_ujian = $selesai_ujian->toDateTimeString();
      $jadwalUjian->lama_ujian = $r->lama_ujian;
      $jadwalUjian->sesi_ujian = $r->sesi_ujian;
      $jadwalUjian->ruang_ujian = $r->ruang_ujian;
      $jadwalUjian->acak_soal = $r->acak_soal;
      $jadwalUjian->pin = $this->generatePin($r->pin_digit);
      $jadwalUjian->tampil_nilai = $r->tampil_nilai;

      if ($jadwalUjian->save()) {
        return redirect()->back()->with('message', 'Data berhasil disimpan');
      }
      return redirect()->back()->withErrors('Terjadi Kesalahan!');
    }

    public function generatePin($digit=4)
    {
        if ($digit<1) {
          $ditig = 4;
        }
        $pin = strtoupper(Str::random($digit));
        $cek = JadwalUjian::where('pin',$pin)->count();
        if ($cek) {
          return $this->generatePin($digit);
        }
        return $pin;
    }

    public function activate($uuid)
    {
      $jadwal = JadwalUjian::where('uuid',$uuid)->first();
      if ($jadwal->aktif) {
        $jadwal->aktif = 0;
        $jadwal->login()->forceDelete();
      }else {
        if (Carbon::parse($jadwal->selesai_ujian) < Carbon::now()) {
          return redirect()->back()->withErrors('Jadwal tidak dapat diaktifkan karena waktu ujian telah berakhir pada '.date('d/m/Y H:i',strtotime($jadwal->selesai_ujian)));
        }
        $jadwal->aktif = 1;
        $jadwal->tes()->update([
          'jawaban'=>null
        ]);
      }
      if ($jadwal->save()) {
        return redirect()->back()->with('message', 'Jadwal ujian '.$jadwal->getSoal->nama.' '.($jadwal->aktif==1?'diaktifkan':'dinonaktifkan'));
      }
      return redirect()->back()->withErrors('Terjadi Kesalahan!');
    }
    public function reset($uuid)
    {
      $jadwal = JadwalUjian::where('uuid',$uuid)->first();
      if ($jadwal->login()->forceDelete()) {
        return redirect()->back()->with('message', 'Jadwal ujian telah direset');
      }
      return redirect()->back()->withErrors('Terjadi Kesalahan!');
    }

    public function edit(Request $r,$uuid)
    {
      if ($r->ajax()) {
        $soal = Soal::all();
        return view("Admin::master.jadwalujian.edit",[
          'soal'=>$soal,
          'data'=>JadwalUjian::where('uuid',$uuid)->first(),
          'kelas'=>Kelas::all()
        ]);
      }
      return redirect()->route('admin.index');
    }

    public function update(Request $r, $uuid)
    {
      if ($r->ajax()) {
        $soal = Soal::where('kode',$r->kode_soal)->first();
        $valid = Validator::make($r->all(),[
          'kode_soal' => 'required',
          'kode_kelas' => 'required',
          'mulai_ujian' => 'required',
          'selesai_ujian' => 'required',
          'lama_ujian' => 'required|numeric|min:1',
          'sesi_ujian' => 'required|numeric|min:1',
        ],[
          'kode_soal.required' => 'Soal ujian tidak boleh kosong',
          'kode_kelas.required' => 'Kelas tidak boleh kosong',
          'mulai_ujian.required' => 'Waktu mulai ujian tidak boleh kosong',
          'mulai_ujian.date_format' => 'Format waktu mulai ujian tidak benar',
          'selesai_ujian.required' => 'Waktu selesai ujian tidak boleh kosong',
          'selesai_ujian.date_format' => 'Format waktu selesai ujian tidak benar',
          'lama_ujian.required' => 'Lama ujian tidak boleh kosong',
          'lama_ujian.numeric' => 'Lama ujian harus berupa angka dalam menit',
          'lama_ujian.min' => 'Lama ujian tidak boleh kurang dari 1 menit',
          'sesi_ujian.required' => 'Sesi ujian tidak boleh kosong',
          'sesi_ujian.numeric' => 'Sesi ujian harus berupa angka',
          'sesi_ujian.min' => 'Sesi ujian tidak boleh kurang dari 1',
          ]);

          if ($valid->fails()) {
            $errs = $valid->errors()->all();
          }else{
            $errs = [];
            if ($r->mulai_ujian!=''&&$r->selesai_ujian!=''&&$r->lama_ujian!='') {
              $now = Carbon::parse(date('Y-m-d H:i'));
              $mulai_ujian = Carbon::createFromFormat('d/m/Y - H:i',$r->mulai_ujian);
              $selesai_ujian = Carbon::createFromFormat('d/m/Y - H:i',$r->selesai_ujian);
              if ($selesai_ujian < $now) {
                array_push($errs,'Waktu selesai ujian tidak boleh kurang dari waktu sekarang');
              }elseif ($mulai_ujian > $selesai_ujian) {
                array_push($errs,'Waktu selesai ujian tidak boleh kurang dari waktu mulai ujian');
              }elseif ($mulai_ujian->addMinutes($r->lama_ujian) > $selesai_ujian) {
                array_push($errs,'Lama ujian melebihi rentang waktu ujian');
              }
            }
          }
          if (count($errs)) {
            return response()->json([
              'success'=>false,
              'messages'=>$errs
            ]);
          }
          return response()->json([
            'success'=>true
          ]);
      }

      $mulai_ujian = Carbon::createFromFormat('d/m/Y - H:i',$r->mulai_ujian);
      $selesai_ujian = Carbon::createFromFormat('d/m/Y - H:i',$r->selesai_ujian);

      $jadwalUjian = JadwalUjian::where('uuid',$uuid)->first();
      $jadwalUjian->kode_soal = $r->kode_soal;
      $jadwalUjian->kode_kelas = $r->kode_kelas;
      $jadwalUjian->mulai_ujian = $mulai_ujian->toDateTimeString();
      $jadwalUjian->selesai_ujian = $selesai_ujian->toDateTimeString();
      $jadwalUjian->lama_ujian = $r->lama_ujian;
      $jadwalUjian->sesi_ujian = $r->sesi_ujian;
      $jadwalUjian->ruang_ujian = $r->ruang_ujian;
      $jadwalUjian->acak_soal = $r->acak_soal;
      $jadwalUjian->tampil_nilai = $r->tampil_nilai;

      if ($jadwalUjian->save()) {
        return redirect()->back()->with('message', 'Data berhasil disimpan');
      }
      return redirect()->back()->withErrors('Terjadi Kesalahan!');
    }

    public function destroy($uuid)
    {
        $jadwalUjian = JadwalUjian::where('uuid',$uuid)->first();
        $jadwalUjian->login()->forceDelete();
        Auth::guard('siswa')->logout();
        $jadwalUjian->tes()->forceDelete();
        if ($jadwalUjian->delete()) {
          return redirect()->back()->with('message', 'Data berhasil dihapus');
        }
        return redirect()->back()->withErrors('Terjadi Kesalahan!');
    }

    public function monitoring(Request $r)
    {
      $jadwalUjian = JadwalUjian::when($r->cari,function($q,$role){
        $role = '%'.$role.'%';
        $q->where('kode','ilike',$role)
        ->orWhere('nama','ilike',$role)
        ->orWhereHas('getSoal',function($q) use($role){
          $q->where('nama','ilike',$role);
        });
      })
      ->where('aktif',1)
      ->paginate(10)->appends(request()->except('page'));
      $data = [
        'title' => 'Monitoring Ujian - Administrator',
        'breadcrumb' => 'Monitoring Ujian',
        'jadwal' => $jadwalUjian
      ];
      return view("Admin::monitoring.index",$data);
    }

    public function monitoringDetail($uuid)
    {
      $jadwal = JadwalUjian::where('uuid',$uuid)->first();
      $login = $jadwal->login()->withTrashed()->get();
      return view("Admin::monitoring.detail",[
        'jadwal'=>$jadwal,
        'login'=>$login,
        'title' => 'Monitoring Ujian - Administrator',
        'breadcrumb' => 'Monitoring '.$jadwal->getSoal->nama,
      ]);
    }

    public function monitoringGetData(Request $r,$uuid)
    {
      if ($r->ajax()) {
        $jadwal = JadwalUjian::where('uuid',$uuid)->first();
        $login = $jadwal->login()->withTrashed()->get();
        return view("Admin::monitoring.getdata",[
          'data'=>$login
        ]);
      }
      return redirect()->route('admin.index');
    }

    public function monitoringStop($pin,$noujian)
    {
      if (request()->ajax()) {
        Login::where('pin',$pin)
        ->where('noujian',$noujian)->update([
          'end'=>Carbon::now()->toDateTimeString()
        ]);
      }
      return redirect()->route('admin.index');
    }
    public function monitoringReset($pin,$noujian)
    {
      if (request()->ajax()) {
        $login = Login::where('pin',$pin)->where('noujian',$noujian)->withTrashed()->first();
        $login->start = null;
        $login->end = null;
        $login->ip_address = null;
        $login->deleted_at = null;
        $login->save();
        // $login->tes()->where('noujian',$noujian)->forceDelete();
        // $login->forceDelete();
      }
      return redirect()->route('admin.index');
    }

    public function print($uuid)
    {
      $jadwal = JadwalUjian::where('uuid',$uuid)->first();
      $kelas = Kelas::where('uuid',$uuid)->first();

      if ($jadwal) {
        $nkelas = $jadwal->kelas?'Kelas '.$jadwal->kelas->nama:'Semua Kelas';
        $filename = 'Kartu Peserta Ujian '.$nkelas.'.pdf';
        $peserta = $jadwal->kelas?$jadwal->kelas->siswa:Siswa::all();
        if (!count($peserta)) {
          return redirect()->back()->withErrors('Data siswa tidak tersedia');
        }
      }elseif ($kelas) {
        if (!count($kelas->siswa)) {
          return redirect()->back()->withErrors('Data siswa tidak tersedia');
        }
        $filename = 'Kartu Peserta Ujian Kelas '.$kelas->nama.' '.$kelas->jurusan.'.pdf';
        $peserta = $kelas->siswa;
      }else {
        $filename = 'Kartu Peserta Ujian.pdf';
        $peserta = Siswa::all();
      }

      // $view = view('Admin::master.jadwalujian.kartu',[
      //   'jadwal'=>$jadwal,
      //   'peserta'=>$peserta,
      //   'title'=>$filename,
      //   'sekolah'=>Sekolah::first(),
      //   'helper'=>new Helper
      // ])->render();
      //
      // $client = new Client;
      // $res = $client->request('POST','http://docker.local:/pdf',[
      //   'form_params'=>[
      //     'html'=>str_replace(url('/'),'http://nginx_ujian/',$view),
      //     'options[page-width]'=>'21.5cm',
      //     'options[page-height]'=>'33cm',
      //     'options[margin-top]'=>'0.5cm',
      //     'options[margin-bottom]'=>'0',
      //     'options[margin-left]'=>'0',
      //     'options[margin-right]'=>'0',
      //   ]
      // ]);
      //
      // if ($res->getStatusCode() == 200) {
      //   return response()->attachment($res->getBody()->getContents(),$filename,'application/pdf');
      // }
      //
      // return redirect()->back()->withErrors(['Tidak dapat mendownload file! Silahkan hubungi operator']);

      $pdf = PDF::loadView('Admin::master.jadwalujian.kartu',[
        'jadwal'=>$jadwal,
        'peserta'=>$peserta,
        'title'=>$filename,
        'sekolah'=>Sekolah::first(),
        'helper'=>new Helper
      ]);

      return $pdf->setOptions([
        'page-width'=>'21.5cm',
        'page-height'=>'33cm'
      ])
      ->setOption('margin-top','0.5cm')
      ->setOption('margin-bottom',0)
      ->setOption('margin-left',0)
      ->setOption('margin-right',0)
      ->stream($filename);
    }
}
