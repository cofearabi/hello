<?php
function dir_file($path){
    @$dir = opendir($path) or exit("can not open directory");
    while($f = readdir($dir))  echo $f . "\n";
    closedir($dir);
}

echo "input folder path";
$a = trim(fgets(STDIN));
dir_file($a);


?>

