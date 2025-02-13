@extends('Admin::layout')
@section('content')
<div class="row">
  <div class="col-md-12">
    <div class="card">
      <div class="card-header card-header-primary">
        <div class="pull-left">
          <h4 class="card-title ">Jadwal Ujian</h4>
          <p class="card-category">Menambah, Mengubah, dan Menghapus Jadwal Ujian</p>
        </div>
        <div class="pull-right">
          <a href="#" class="btn btn-sm btn-warning" data-toggle="modal" data-backdrop="static" data-keyboard="false" data-target="#modalEdit" data-url="{{ route('jadwal.ujian.create') }}"><i class="material-icons">add</i> Tambah Jadwal</a>
        </div>
      </div>
      <div class="card-body">
        <div class="table-responsive">
          <table class="table table-hover">
            <thead class=" text-primary">
              <th>No</th>
              <th>Soal</th>
              <th>Kelas</th>
              <th>Mulai</th>
              <th>Selesai</th>
              <th>Lama Ujian</th>
              <th>Sesi</th>
              <th>PIN</th>
              <th></th>
            </thead>
            <tbody>
              @if (count($jadwal))
                @foreach ($jadwal as $key => $v)
                  <tr>
                    @php
                    $index = Request::get('page')??1;
                    @endphp
                    <td style="vertical-align: top">{{ (($index-1)*10)+$key+1 }}</td>
                    <td style="vertical-align: top">{{ $v->getSoal->nama?'('.$v->getSoal->kode.') '.$v->getSoal->nama:'-' }}</td>
                    <td style="vertical-align: top">{{ $v->kelas?'('.$v->kelas->kode.') '.$v->kelas->nama.' '.$v->kelas->jurusan:'Semua' }}</td>
                    <td style="vertical-align: top">{{ date('d/m/Y H:i',strtotime($v->mulai_ujian)) }}</td>
                    <td style="vertical-align: top">{{ date('d/m/Y H:i',strtotime($v->selesai_ujian)) }}</td>
                    <td style="vertical-align: top">{{ $v->lama_ujian.' Menit' }}</td>
                    <td style="vertical-align: top">{{ $v->sesi_ujian }}</td>
                    <td style="font-weight:bold;vertical-align: top" class="text-primary">{{ $v->pin }}</td>
                    <td style="white-space: nowrap;width: 50px;vertical-align: top" class="text-right">
                      @if ($v->getSoal->item->count()&&($v->aktif||is_null($v->aktif)))
                      <a class="btn btn-sm btn-xs {{ $v->aktif?'btn-warning':'btn-primary' }} confirm" title="{{ $v->aktif?'Nonaktifkan':'Aktifkan' }} Jadwal Ujian" data-text="{{ $v->aktif?'Nonaktifkan Jadwal Ujian?<br>Semua peserta akan logout!':'Hasil ujian sebelumnya akan terhapus!<br>Aktifkan Jadwal Ujian?' }}" href="#" data-url="{{ route('jadwal.ujian.activate',['uuid'=>$v->uuid]) }}">{!! $v->aktif?'Nonaktifkan':'<i class="material-icons">check</i>' !!}</a>
                      @if ($v->aktif)
                        <a class="btn btn-sm btn-xs btn-info" title="Lihat status peserta" href="{{ route('jadwal.ujian.monitoring.detail',['uuid'=>$v->uuid]) }}" class="text-info"><i class="material-icons">desktop_windows</i></a>
                      @endif
                      @endif
                      @if (($v->kelas&&count($v->kelas->siswa))||$v->kode_kelas=='all')
                        <a href="{{ route('jadwal.ujian.print',['uuid'=>$v->uuid]) }}" target="_blank" class="btn btn-sm btn-xs btn-success" title="Cetak Kartu Ujian"><i class="material-icons">print</i></a>
                      @endif
                      @if (!$v->aktif)
                      @if ($v->aktif||is_null($v->aktif))
                      <a class="btn btn-sm btn-xs btn-info" title="Ubah" href="#" data-toggle="modal" data-backdrop="static" data-keyboard="false" data-target="#modalEdit" data-url="{{ route('jadwal.ujian.edit',['uuid'=>$v->uuid]) }}" class="text-info"><i class="material-icons">edit</i></a>
                      @endif
                      <a class="btn btn-sm btn-xs btn-danger confirm" title="Hapus" data-text="Semua jawaban & nilai peserta akan terhapus untuk jadwal ini! Hapus Jadwl Ujian {{ $v->getSoal->nama.' ('.$v->getSoal->kode.')' }}" href="#" data-url="{{ route('jadwal.ujian.destroy',['uuid'=>$v->uuid]) }}"><i class="material-icons">delete</i></a>
                      @endif
                    </td>
                  </tr>
                @endforeach
              @else
                <tr>
                  <td class="text-center no-data">Data tidak tersedia</td>
                </tr>
              @endif
            </tbody>
          </table>
          {{ $jadwal->links() }}
        </div>
      </div>
    </div>
  </div>
</div>
@endsection
@section('footer')
  <div class="modal fade" id="modalEdit" role="dialog" aria-labelledby="" aria-hidden="true">
    <div class="modal-dialog">
      <div class="modal-content"></div>
    </div>
  </div>
  <script type="text/javascript">
  $(document).ready(function(){
    $('#modalEdit').on('show.bs.modal', function (e) {
      var _this = $(this);
      var _data = $(e.relatedTarget);
      _this.find('.modal-dialog').css('max-width','200px')
      _this.find('.modal-content').html('<h4 style="margin: 0">Silahkan tunggu...</h4>');
      $.get(_data.data('url'),{},function(res){
        _this.find('.modal-dialog').animate({'max-width':'400px'},150,function(){
          _this.find('.modal-content').html(res)
        })
      });
    });
  })
  @if (session()->has('message'))
    md.showNotification('bottom','right','{{ session()->get('message') }}','success','check');
  @endif
  @if ($errors->any())
    @foreach ($errors->all() as $error)
      md.showNotification('bottom','right','{{ $error }}','danger','not_interested');
    @endforeach
  @endif
  </script>
@endsection
