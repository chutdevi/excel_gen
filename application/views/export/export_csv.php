<?php

date_default_timezone_set("Asia/Bangkok");
            //echo $csv; exit;

            $str = $csv . "\r\n" ;

            foreach ($list_act_report as $row => $value) 
            {
                foreach ( $value as $r => $val)
                {
                    // if ($val=='L40810') 
                    // {
                    //     $val = 'M50040';
                    // }
                   $val  =  '"'.str_replace(",", ";", $val ).'"';
                   $str .=  $val. ","; 
                }
                $str = substr($str, 0, (strlen($str)-1) )  . "\r\n" ;  
            }
	$str = substr($str, 0, (strlen($str)-2) ) ; 
    $myfile = fopen("filedownload/purc/".$filename, 'w') or die("Unable to open file!");
    

    ///file_get_contents('http://192.168.161.102/dep_trainer/Api_tool/api_test');
    //echo $str; exit;
    fwrite($myfile,$str);
    fclose($myfile);

    //echo $sta; exit;
    if($sta > 10)
    {
        rename("filedownload/purc/".$filename, "filedownload/purc_his/".$filename);
        output_file("filedownload/purc_his/".$filename);
    }
    else
        output_file("filedownload/purc/".$filename); 
    



function output_file($namefile){
        //$namefile = "Query_pods_data.sql";
        $file = $namefile; 
        //echo basename($file); exit;
        header("Content-Description: File Transfer"); 
        header("Content-Type: application/octet-stream"); 
        header("Content-Disposition: attachment; filename=".basename($file) ); 
        readfile ($file);
        exit;
}

?>

 
