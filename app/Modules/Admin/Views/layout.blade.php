<!doctype html>
<html lang="en">

<head>
  <title>{{ $title??'Administrator' }}</title>
  <!-- Required meta tags -->
  <meta charset="utf-8">
  <meta content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0" name="viewport" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
  @php
    $sekolah = App\Models\Sekolah::first();
  @endphp
  @if ($sekolah&&is_file(base_path('uploads/'.$sekolah->logo)))
    <link rel="icon" href="{{ url('uploads/'.$sekolah->logo) }}" type="image/x-icon"/>
    <link rel="shortcut icon" href="{{ url('uploads/'.$sekolah->logo) }}" type="image/x-icon"/>
  @endif
  <!--     Fonts and icons     -->
  <link href="{{ url('assets/css/fonts.css') }}" rel="stylesheet" />
  <link href="{{ url('assets/css/font-awesome.min.css') }}" rel="stylesheet" />
  <!-- Material Kit CSS -->
  <link href="{{ url('assets/iconfont/material-icons.css') }}" rel="stylesheet" />
  <link href="{{ url('assets/css/select2.min.css') }}" rel="stylesheet" />
  <link href="{{ url('assets/css/bootstrap-material-datetimepicker.css') }}" rel="stylesheet" />
  <link href="{{ url('assets/css/material-dashboard.css?v=2.1.1') }}" rel="stylesheet" />
  <style media="screen">
    .pagination{
      float: right;
    }
    .pagination .page-link{
      color: #9c27b0
    }
    .pagination .page-item.active .page-link{
      background: #9c27b0;
    }
    .btn-xs{
      padding: 0.4rem !important;
    }
    .table td .btn-xs{
      margin: 0 !important;
    }
    select.form-control {
        /* width: 268px; */
        /* padding: 5px;
        font-size: 16px;
        line-height: 1;
        border: 0;
        border-radius: 5px;
        height: 34px; */
        background: url({{ url('assets/img/br_down.png') }}) no-repeat right !important;
        /* -webkit-appearance: none; */
        border-bottom: solid 1px #999999;
        background-position-x: 240px;
        background-size: 11px !important;
    }
    select.form-control:focus{
      border-bottom: solid 2px #9c27b0;
    }
    .alert.animated{
      z-index: 1051 !important;
    }
    .text-bold{
      font-weight: bold !important;
    }
    .select2-results__option{
      display: block !important;
      text-align: left !important;
    }
    .select2-container{
      width: 100% !important;
    }
    button.btn.btn-flat {
        padding: 1rem;
        background: none;
        color: #333;
        border: none !important;
    }
    .select2-container .select2-selection--single{
      height: 35px;
    }
    .select2-container--default .select2-selection--single .select2-selection__rendered{
      line-height: 34px;
    }
    .select2-container--default .select2-selection--single .select2-selection__arrow{
      height: 34px;
    }
    .modal-content{
      padding: 0 15px;
    }
  </style>
  @yield('header')
</head>

<body>
  <div class="wrapper ">
    @include('Admin::sidebar')
    <div class="main-panel">
      <!-- Navbar -->
      <nav class="navbar navbar-expand-lg navbar-transparent navbar-absolute fixed-top ">
        <div class="container-fluid">
          <div class="navbar-wrapper">
            <a class="navbar-brand" href="javascript:void(0)">{{ $breadcrumb??'Dashboard' }}</a>
          </div>
          <button class="navbar-toggler" type="button" data-toggle="collapse" aria-controls="navigation-index" aria-expanded="false" aria-label="Toggle navigation">
            <span class="sr-only">Toggle navigation</span>
            <span class="navbar-toggler-icon icon-bar"></span>
            <span class="navbar-toggler-icon icon-bar"></span>
            <span class="navbar-toggler-icon icon-bar"></span>
          </button>
          <div class="collapse navbar-collapse justify-content-end">
            <form class="navbar-form" id="search-bar" action="{{ url()->current() }}">
              <div class="input-group no-border">
                <input type="text" name="cari" value="{{ Request::get('cari') }}" class="form-control" placeholder="Cari...">
                <button type="submit" class="btn btn-white btn-round btn-just-icon">
                  <i class="material-icons">search</i>
                  <div class="ripple-container"></div>
                </button>
              </div>
            </form>
            <ul class="navbar-nav">
              <li class="nav-item dropdown">
                <a class="nav-link" href="#pablo" id="navbarDropdownProfile" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
                  <i class="material-icons">person</i>
                  <p class="d-lg-none d-md-block">
                    Account
                  </p>
                </a>
                <div class="dropdown-menu dropdown-menu-right" aria-labelledby="navbarDropdownProfile">
                  <a class="dropdown-item" href="{{ route('admin.profile') }}">Edit Profil</a>
                  <div class="dropdown-divider"></div>
                  <a class="dropdown-item" href="{{ route('admin.logout') }}">Log out</a>
                </div>
              </li>
              <!-- your navbar here -->
            </ul>
          </div>
        </div>
      </nav>
      <!-- End Navbar -->
      <div class="content">
        <div class="container-fluid">
          @yield('content')
        </div>
      </div>
      <footer class="footer">
        <div class="container-fluid">
          <div class="copyright float-right">
            &copy; {{ date('Y') }} Aplikasi Ujian by <a href="https://www.facebook.com/aezdar"> Asdar Bin Syam </a>
          </div>
          <script src="{{ url('/') }}/assets/js/core/jquery.min.js"></script>
          <script src="{{ url('/') }}/assets/js/core/popper.min.js"></script>
          <script src="{{ url('/') }}/assets/js/core/bootstrap-material-design.min.js"></script>
          <script src="{{ url('/') }}/assets/js/plugins/perfect-scrollbar.jquery.min.js"></script>
          <script src="{{ url('/') }}/assets/js/plugins/bootstrap-notify.js"></script>
          <script src="{{ url('/') }}/assets/js/material-dashboard.js?v=2.1.1" type="text/javascript"></script>
          <script src="{{ url('/') }}/assets/js/select2.min.js" type="text/javascript"></script>
          <script src="{{ url('/') }}/assets/js/moment.min.js" type="text/javascript"></script>
          <script src="{{ url('/') }}/assets/js/bootstrap-material-datetimepicker.js" type="text/javascript"></script>
          <script type="text/javascript">
          $(document).ready(function(){
            $.fn.modal.Constructor.prototype._enforceFocus = function() {};
            $(".modal").on('show.bs.modal',function(){
              $(".content").css({'position':'fixed','width':$(".content").width()+30});
              $(".copyright").hide();
            })
            $(".modal").on('hide.bs.modal',function(){
              $(".content").css({'position':'','width':''});
              $(".copyright").show();
            })
            $(".no-data").each(function(){
              $(this).attr('colspan',$(this).closest('table').find('th').length)
            })
            $(".delete").click(function(e){
              e.preventDefault();
              var confirm = `
                <h3 class="text-center">Yakin ingin menghapus `+$(this).data('text')+`?</h3>
                <div class="text-center" style="margin-bottom: 15px">
                  <button type="button" class="btn btn-default" data-dismiss="modal">Tidak</button>
                  <a href="`+$(this).data('url')+`" class="btn btn-danger">Ya</a>
                </div>
              `;
              $("#modalDelete").find('.modal-content').html(confirm);
              $("#modalDelete").modal({
                show: true,
                // keyboard: false,
                backdrop: 'static'
              });
            })
            $(".confirm").on('click',function(e){
              e.preventDefault();
              var confirm = `
                <h3 class="text-center">`+$(this).data('text')+`</h3>
                <div class="text-center" style="margin-bottom: 15px">
                  <button type="button" class="btn btn-default" data-dismiss="modal">Tidak</button>
                  <a href="`+$(this).data('url')+`" class="btn btn-danger">Ya</a>
                </div>
              `;
              $(".modal-confirm").find('.modal-content').html(confirm);
              $(".modal-confirm").modal({
                show: true,
                // keyboard: false,
                backdrop: 'static'
              });
            })
          })
          $(".notif").click(function(e){
            e.preventDefault();
            var confirm = `
              <h3 class="text-center">`+$(this).data('text')+`</h3>
              <div class="text-center" style="margin-bottom: 15px">
                <button type="button" class="btn btn-primary btn-sm" data-dismiss="modal">Tutup</button>
              </div>
            `;
            $(".modal-confirm").find('.modal-content').html(confirm);
            $(".modal-confirm").modal({
              show: true,
              // keyboard: false,
              backdrop: 'static'
            });
          })
          $(document).ajaxSuccess(function() {
            setTimeout(()=>{
              $("select").select2();
              $('#dend').bootstrapMaterialDatePicker({ format : 'DD/MM/YYYY - HH:mm',minDate: '{{ Carbon\Carbon::now()->format('D M d Y H:i:s e+') }}'}).change(function(e,date){
                if ($('#dstart').val()=='') {
                  $('#dstart').bootstrapMaterialDatePicker('setDate',date);
                }
              })
              $('#dstart').bootstrapMaterialDatePicker({ format : 'DD/MM/YYYY - HH:mm'}).change(function(e,date){
                $('#dend').bootstrapMaterialDatePicker('setDate',date);
                $('#dend').bootstrapMaterialDatePicker({minDate:date});
              });
            },175)
          })
          </script>
          <div class="modal fade modal-confirm" id="modalDelete" tabindex="-1" role="dialog" aria-labelledby="" aria-hidden="true">
            <div class="modal-dialog">
              <div class="modal-content"></div>
            </div>
          </div>
          @yield('footer')
          <!-- your footer here -->
        </div>
      </footer>
    </div>
  </div>
</body>

</html>
