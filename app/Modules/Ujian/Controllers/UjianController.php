<?php

namespace App\Modules\Ujian\Controllers;

use Illuminate\Http\Request;
use App\Http\Controllers\Controller;
use App\Models\Sekolah;
use App\Models\Siswa;
use App\Models\JadwalUjian;
use App\Models\Login;
use App\Models\ItemSoal;
use App\Models\Tes;
use App\Models\Soal;
use Illuminate\Support\Str;
use Validator;
use Auth;
use Hash;
use Carbon\Carbon;

class UjianController extends Controller
{

    public function index()
    {
      return view("Ujian::index",[
        'sekolah'=>Sekolah::first()
      ]);
    }

    public function login(Request $r)
    {
      Validator::make($r->all(),[
        'noujian'=>'required',
        'password'=>'required',
        'pin'=>'required',
      ],[
        'noujian.required'=>'No ujian tidak boleh kosong',
        'password.required'=>'Password tidak boleh kosong',
        'pin.required'=>'Pin harus diisi',
      ])->validate();

      $siswa = Siswa::where('noujian',$r->noujian)->first();

      if (Hash::check($r->password,$siswa->password)) {
        $jadwal = JadwalUjian::where('pin',$r->pin)->where('aktif',1)->first();
        if (!$jadwal) {
          Auth::guard('siswa')->logout();
          return redirect()->route('ujian.login')->withErrors(['Ujian tidak tersedia'])->withInput($r->only('noujian'));
        }elseif (Carbon::now() <= Carbon::parse($jadwal->mulai_ujian)) {
          Auth::guard('siswa')->logout();
          return redirect()->route('ujian.login')->withErrors(['Sesi ujian akan dimulai pada tanggal '.date('d/m/Y',strtotime($jadwal->mulai_ujian)).' pukul '.date('H:i',strtotime($jadwal->mulai_ujian))])->withInput($r->only('noujian'));
        }elseif (Carbon::now() >= Carbon::parse($jadwal->selesai_ujian)) {
          Auth::guard('siswa')->logout();
          return redirect()->route('ujian.login')->withErrors(['Sesi ujian telah berakhir'])->withInput($r->only('noujian'));
        }elseif (!in_array($siswa->uuid,json_decode($jadwal->peserta))) {
          Auth::guard('siswa')->logout();
          return redirect()->route('ujian.login')->withErrors(['Anda tidak dapat mengikuti ujian ini'])->withInput($r->only('noujian'));
        }elseif ($siswa->attemptLogin()->where('pin',$r->pin)->first()) {
          if (!is_null($siswa->attemptLogin()->where('pin',$r->pin)->first()->end)) {
            Auth::guard('siswa')->logout();
            return redirect()->route('ujian.login')->withErrors(['Anda sudah selesai ujian'])->withInput($r->only('noujian'));
          }elseif (!is_null($siswa->attemptLogin()->where('pin',$r->pin)->first()->_token)) {
            Auth::guard('siswa')->logout();
            return redirect()->route('ujian.login')->withErrors(['Anda sudah login di tempat lain!'])->withInput($r->only('noujian'));
          }
          Auth::guard('siswa')->attempt([
            'noujian'=>$r->noujian,
            'password'=>$r->password
          ],1);
          $user = Auth::guard('siswa')->user();
          $ujian = Auth::guard('siswa')->user()->attemptLogin()->where('pin',$r->pin)->first();
          $ujian->_token = $user->remember_token;
          $ujian->ip_address = $r->ip();
          $ujian->save();
          session()->flush();
          return redirect()->back();
        }
        Auth::guard('siswa')->attempt([
          'noujian'=>$r->noujian,
          'password'=>$r->password
        ],1);
        $user = Auth::guard('siswa')->user();
        $ujian = new Login;
        $ujian->uuid = (string)Str::uuid();
        $ujian->noujian = $user->noujian;
        $ujian->_token = $user->remember_token;
        $ujian->pin = $r->pin;
        $ujian->start = null;
        $ujian->end = null;
        $ujian->ip_address = $r->ip();
        $ujian->save();
        session()->flush();
        return redirect()->back();
      }

      return redirect()->back()->withErrors(['Data login tidak benar!'])->withInput($r->only('noujian','pin'));
    }

    public function cekData()
    {
      return view('Ujian::cekdata',[
        'title'=>'Pemeriksaan Data Ujian',
        'breadcrumb'=>'Pemeriksaan Data Ujian',
        'siswa'=>Auth::guard('siswa')->user()
      ]);
    }

    public function timer($int=false)
    {
      $login = Auth::guard('siswa')->user()->login;
      $jadwal = $login->jadwal;
      $timerNow = Carbon::now()->addMinutes($jadwal->lama_ujian) <= Carbon::parse($jadwal->selesai_ujian) ? Carbon::now()->addMinutes($jadwal->lama_ujian) : Carbon::parse($jadwal->selesai_ujian);

      $intval = $timerNow->diffInSeconds(Carbon::parse($login->start)->addMinutes($jadwal->lama_ujian));

      if ($int) {
        return $intval;
      }

      return $timerNow->subSeconds($intval);
    }

    public function tes()
    {
      $login = Auth::guard('siswa')->user()->login;
      $jadwal = $login->jadwal;
      if ($login->start==null) {
        $soal = [];
        $login->start = Carbon::now()->toDateTimeString();

        $soalItem = ItemSoal::where('jenis_soal',$jadwal->jenis_soal)
        ->whereHas('getSoal',function($q) use($jadwal){
          $q->whereIn('uuid',json_decode($jadwal->soal));
        })
        ->select('uuid');
        if ($jadwal->acak_soal=='Y') {
          $soalItem = $soalItem->inRandomOrder();
        }
        $soalItem = $soalItem->get();
        $i = 0;
        if (count($soalItem)) {
          foreach ($soalItem as $key => $si) {
            $i++;
            if ($i > $jadwal->jumlah_soal) {
              break;
            }
            array_push($soal,$si->uuid);
          }
        }
        $login->soal_ujian = json_encode($soal);
        $login->save();
      }

      $timer = $this->timer();

      return view('Ujian::tes',[
        'title'=>'Mulai Ujian '.$jadwal->nama_ujian,
        'breadcrumb'=>'Ujian '.$jadwal->nama_ujian,
        'soal'=>$login->soal_ujian,
        'now'=>Carbon::now(),
        'timer'=>$timer,
        'siswa'=>Auth::guard('siswa')->user()
      ]);
    }

    public function getsoal(Request $r)
    {
      if ($r->ajax()) {
        $siswa = Auth::guard('siswa')->user();

        if (!$siswa || !$siswa->login || @$siswa->login->end) {
          return $err = 0;
        }

        if ($r->has('checking')) {
          sleep(3);
          return response()->json([
            'status' => 1,
            'token' => csrf_token()
          ]);
        }

        $current_number  = array_search($r->soal,json_decode($siswa->login->soal_ujian));

        $siswa->login->update(['current_number'=>$current_number]);

        $soal = ItemSoal::where('uuid',$r->soal)->first();

        $opsis = null;

        $cek = Tes::where('noujian',$siswa->noujian)
        ->where('pin',$siswa->login->pin)
        ->where('soal_item',$soal->uuid)
        ->first();

        if (!$cek) {
          if ($soal->jenis_soal=='P') {
            if ($soal->acak_opsi=='Y') {
              $keys = array_keys(json_decode($soal->opsi));
              shuffle($keys);
              foreach ($keys as $v) {
                $opsis[$v] = json_decode($soal->opsi)[$v];
              }
            }else {
              $opsis = json_decode($soal->opsi);
            }
          }

          $tes = new Tes;
          $tes->noujian = $siswa->noujian;
          $tes->kode_soal = $soal->kode_soal;
          $tes->pin = $siswa->login->pin;
          $tes->soal_item = $soal->uuid;
          $tes->opsi = json_encode($opsis);

          $tes->save();

        }else {
          $opsis = json_decode($cek->opsi);
        }

        return view('Ujian::soal',[
          'soal'=>$soal,
          'siswa'=>$siswa,
          'allSoal'=>$siswa->login->soal_ujian,
          'opsis'=>$opsis,
          'key'=>$r->key
        ]);
      }
      return redirect()->route('ujian.login');
    }

    public function submit(Request $r,$uuid)
    {
      if ($r->ajax()) {
        $siswa = Auth::guard('siswa')->user();
        $soal = ItemSoal::where('uuid',$uuid)->first();
        $tes = Tes::where('noujian',$siswa->noujian)
        ->where('pin',$siswa->login->pin)
        ->where('soal_item',$soal->uuid)
        ->first();

        $tes->jawaban = $r->jawab;
        if ($tes->save()) {
          return response()->json(['success'=>true]);
        }
        return response()->json(['success'=>false]);
      }
      return redirect()->route('ujian.login');
    }

    public function selesai()
    {
      $siswa = Auth::guard('siswa')->user();
      $siswa->login()->update([
        'end'=>Carbon::now()->toDateTimeString()
      ]);
      return redirect()->route('ujian.nilai');
    }

    public function nilai()
    {
      $siswa = Auth::guard('siswa')->user();
      if (is_null($siswa->login->end)) {
        $siswa->login()->update([
          'end'=>Carbon::now()
        ]);
      }
      $nilai = null;
      $nbenar = null;
      $jadwal = $siswa->login->jadwal;
      $bobot = $jadwal->bobot;
      $jumlah_soal = count(json_decode($siswa->login->soal_ujian));
      $dtes = Tes::where('noujian',$siswa->noujian)
      ->where('pin',$siswa->login->pin)->whereIn('soal_item',json_decode($siswa->login->soal_ujian))->get();
      foreach ($dtes as $key => $tes) {
        $benar = $tes->soalItem->benar;
        if (!is_null($benar) && (string) $tes->jawaban == (string) $benar && $tes->soalItem->jenis_soal=='P') {
          if (is_null($nbenar)) {
            $nbenar = 0;
          }
          $nbenar++;
        }
      }
      if ($jumlah_soal) {
        $nilai = 0;
      }
      if (!is_null($nbenar)) {
        $nilai += round($nbenar/$jumlah_soal*$bobot,2);
      }
      return view('Ujian::nilai',[
        'title'=>'Hasil Ujian',
        'breadcrumb'=>'Hasil Ujian',
        'siswa'=>$siswa,
        'nilai'=>$nilai
      ]);
    }

    public function audioRepeat(Request $r)
    {
      if ($r->ajax()) {
        $soal = $r->soal;
        $audio = $r->audio;
        $count = $r->count;
        $scount = $soal.$audio.'count';

        session()->put([$scount=>$r->count]);
        return session()->get($scount);
      }
      return redirect()->route('ujian.login');
    }

    public function logout()
    {
      $siswa = Auth::guard('siswa');
      if ($siswa->user()) {
        // $siswa->user()->login()->delete();
        $siswa->logout();
        session()->flush();
      }
      return redirect()->route('ujian.login');
    }

}
