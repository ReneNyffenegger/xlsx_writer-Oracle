define step=&1
prompt &step

$&test_out_dir\&step..xlsx
$cmd /c rmdir /s /q gotten\&step
$&unzip_cmd &test_out_dir\&step..xlsx gotten\&step
