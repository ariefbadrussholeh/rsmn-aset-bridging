<?php
error_reporting(E_ALL^(E_NOTICE | E_WARNING | E_DEPRECATED));
$hostname_conn = "192.168.0.233";
$database_conn = "dbaset_rsijs";
$username_conn = "umum";
$password_conn = "RSmndb2020";

$konek = mysql_connect($hostname_conn,$username_conn,$password_conn,$database_conn)or die("Koneksi gagal");
mysql_select_db($database_conn,$konek)or die("Database tidak bisa dibuka");

date_default_timezone_set("Asia/Jakarta");
$bulan = array('','Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember');

$perpage=100;
$def_loc = "/rsmn/aset";
$singkatan = "RSMN";
$cop_rs  = "<table border='0'>
		<tr>
			<td style='font-size: 12px' align='left'>
				<img src='../images/logo_rs.jpg' width='65' height='80'/>
			</td>
			<td style='font-size: 12px' align='left'>
				PEMERINTAH PROVINSI JAWA TIMUR <br> 
				DINAS KESEHATAN <br>
				RUMAH SAKIT UMUM DAERAH MOHAMMAD NOER PAMEKASAN<br>
				Jl. Bonorogo No.17, Kec. Pademawu, Kab. Pamekasan<br>
				Telp. (0324) 322594 - Fax. (0324) 323085<br>
			</td>
		</tr>
	</table>";

function tglSQL($tgl){
   $t=explode("-",$tgl);
   $t=$t[2].'-'.$t[1].'-'.$t[0];
   return $t;
}

function cekNull($val){
    if($val == '' || !isset($val) || $val == null){
        return 0;
    }
    else{
        return $val;
    }
}

function getNamaBulan($ibln,$iMode){
   switch ($ibln){
      case 1: if($iMode==0) return 'Januari'; else return 'Jan'; break;
      case 2: if($iMode==0) return 'Pebruari'; else return 'Peb';break;
      case 3: if($iMode==0) return 'Maret'; else return 'Mar';break;
      case 4: if($iMode==0) return 'April'; else return 'Apr';break;
      case 5: if($iMode==0) return 'Mei'; else return 'Mei';break;
      case 6: if($iMode==0) return 'Juni'; else return 'Jun';break;
      case 7: if($iMode==0) return 'Juli'; else return 'Jul';break;
      case 8: if($iMode==0) return 'Agustus'; else return 'Agust';break;
      case 9: if($iMode==0) return 'September'; else return 'Sept';break;
      case 10: if($iMode==0) return 'Oktober'; else return 'Okt';break;
      case 11: if($iMode==0) return 'Nopember'; else return 'Nop';break;
      case 12: if($iMode==0) return 'Desember'; else return 'Des';break;
   }
}

function currency($angka)
{
	$rupiah= number_format($angka,2,',','.');
	return $rupiah;
}
?>
