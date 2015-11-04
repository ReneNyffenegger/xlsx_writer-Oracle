define step=&1
prompt &step

-- Open created xls
-- $&test_out_dir\&step..xlsx

$cmd /c rmdir /s /q gotten\&step
$&unzip_cmd &test_out_dir\&step..xlsx gotten\&step

-- Diff expected with what was written
$diff -r expected\&step gotten\&step
