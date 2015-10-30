define step=&1
prompt &step

$&test_out_dir\&step..xlsx
$rmdir /s /q gotten\&step
$&unzip_cmd &test_out_dir\&step..xlsx gotten\&step
