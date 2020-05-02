<?php
defined('BASEPATH') OR exit('No direct script access allowed');
 
date_default_timezone_set('Asia/Bangkok');

class Query extends CI_Controller {

	public function __construct()
	{ 
		parent::__construct();



	}

	public function index()
	{


	}

	public function prod_query()
	{





		$this->load->view('Export/download_prod_fa');

	}



	public function td( $dat )
	{
					$content = file_get_contents('http://192.168.161.102/report_access/Api_tool/api_test/'. $dat  );
					// แปลงข้อมูลที่รับมาในรูป json มาเป็น array จะได้ใช้ง่าย ๆ
					$DATA = json_decode($content, true);

					// //dump ข้อมูลออกมาดู
					//print_r($DATA);
					//print_r($DATA);
					// ลองดึงออกทีล่ะค่า
					//echo "<hr>";


      				//exit;


	}

}

?>