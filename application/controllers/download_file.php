<?php
defined('BASEPATH') OR exit('No direct script access allowed');
 
date_default_timezone_set('Asia/Bangkok');

class download_file extends CI_Controller 
{

	public function __construct()
	{ 
		parent::__construct();



	}

	public function index( $filename )
	{


		//echo str_replace('-','/', $filename); exit;
		$data['filename'] = str_replace('-','/', $filename);

		$this->load->view('Export/download_csv',$data);

	}


}

?>