<?php

function copyFile($f){
     $path = getcwd() . "/backup";
    if(!file_exists($path)) mkdir($path);
    copy(getcwd() . "/" . $f,$path . "/" . $f);
}

echo "ファイル名を入力:";
$a = trim(fgets(STDIN));
copyFile($a);

?>

