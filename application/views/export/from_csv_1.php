<?php

date_default_timezone_set("Asia/Bangkok");
            //echo $csv; exit;
            unset($str);
            $str = $csv ."\r\n";

// echo "<pre>";
// echo "start";
// echo $str; 
//echo $str; exit;
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
	$str = substr($str, 0, (strlen($str)-2) );
    $myfile = fopen("filedownload/sale/".$filename, 'w+') or die("Unable to open file!");
    fwrite($myfile, $str);
    fclose($myfile);
    
    if($sta > 10)
    {
    rename("filedownload/sale/".$filename, "filedownload/sale_his/".$filename);

    output_file("filedownload/sale_his/".$filename);
    }
    else
    output_file("filedownload/sale/".$filename);



function output_file($namefile){
        //$namefile = "Query_pods_data.sql";
        $file = $namefile; 
        header("Content-Description: File Transfer"); 
        header("Content-Type: application/octet-stream"); 
        header("Content-Disposition: attachment; filename=" . basename($file) ); 
        readfile ($file);
        exit;
}

?>

 
