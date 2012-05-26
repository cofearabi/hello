#!/usr/bin/perl
use strict;
use warnings;
use Web::Scraper;
use URI;
use Encode;
use Spreadsheet::WriteExcel;
use strict;
use warnings;
use utf8;
# binmode STDIN, ':encoding(cp932)';
# binmode STDOUT, ':encoding(cp932)';
# binmode STDERR, ':encoding(cp932)';

my $time = time();

# my $thisDay = thisDay($time);
# my $prev50Day = prev50Day($time,12);

# my @start;
# my @end;

my @hajimari;
my @owari;

 $hajimari[5] = prev50Day($time,366);
 $owari[5] = prev50Day($time,306);
 $hajimari[4] = prev50Day($time,305);
 $owari[4] = prev50Day($time,245);
 $hajimari[3] = prev50Day($time,244);
 $owari[3] = prev50Day($time,184);
 $hajimari[2] = prev50Day($time,183);
 $owari[2] = prev50Day($time,123);
 $hajimari[1] = prev50Day($time,122);
 $owari[1] = prev50Day($time,62);
 $hajimari[0] = prev50Day($time,61);
 $owari[0]    = thisDay($time);
# my $start[0] = prev50day($time,185);
# my $end[0]   = prev50day($time,124);
# my $start[1] = prev50day($time,123);
# my $end[1]   = prev50day($time,62);
# my $start[2] = prev50day($time,61);
# my $end[2]   = thisday($time)

sub thisDay{
  my $time = shift || time();
  my $thisDay = $time ;
  my ($yyyy, $mm, $dd) = (localtime($thisDay))[5,4,3];

  $yyyy += 1900;
  $mm += 1;

  return(
    sprintf('%4d/%02d/%02d', $yyyy, $mm, $dd)
  );
}
sub prev50Day{
  my $time = shift || time();
  my $prevDay = $time - (24 * 60 * 60 * $_[0]);
  my ($yyyy, $mm, $dd) = (localtime($prevDay))[5,4,3];

  $yyyy += 1900;
  $mm += 1;

  return(
    sprintf('%4d/%02d/%02d', $yyyy, $mm, $dd)
  );
}

    # 新しいExcelワークブックの作成
    my $workbook = Spreadsheet::WriteExcel->new("c:\\temp.xls");


# my @list = split(/\//, $thisDay );
# print $list[0]."\n";
# print $list[1]."\n";
# print $list[2]."\n";
#   print $list[3]."\n";

# my @list1 = split(/\//, $prev50Day );
# print $list1[0]."\n";
# print $list1[1]."\n";
# print $list1[2]."\n";

    # 新しいExcelワークブックの作成
    # my $workbook = Spreadsheet::WriteExcel->new("temp.xls");




open(IN,"zzz.txt");
while( my $Meigara = <IN> )
{
sleep(0.5);
chomp($Meigara);

    # ワークシートの追加
    my $worksheet = $workbook->addworksheet($Meigara);
    $worksheet->set_column('B:B', 18);
    $worksheet->set_column('A:A', 18);
    $worksheet->set_column('G:G', 12);

    #  書式の追加と定義
#    $format = $workbook->addformat(); # 書式の追加
#    $format->set_bold();
#a3    $format->set_color('red');
#    $format->set_align('center');

    # 行、列の書き方で書式付とそうでない文字列を出力
#    $col = $row = 0;
#    $worksheet->write($row, $col, "Hi Excel!", $format);
#    $worksheet->write(1,    $col, "Hi Excel!");

    # A1という書き方を使って、数字と式を出力
#    $worksheet->write('A3', 1.2345);
#    $worksheet->write('A4', '=SIN(PI()/4)');

my $ws = scraper {
  process
    '//center/div[@class="invest"]/table[2]/tr[2]/td[3]',
    meigara => [ 'text', sub { s/,//g } ];
  process
    '//center/div[@class="invest"]/table[2]/tr[2]/td[1]',
    code => [ 'text', sub { s/,//g } ];
  process
    '//center/div[@class="invest"]/table[2]/tr[2]/td[5]/b',
    price => [ 'text', sub { s/,//g } ];
  process
    '//center/div[@class="invest"]/table[2]/tr[2]/td[4]',
    jikoku=> [ 'text', sub { s/,//g } ];
};

my $price = 0;
my $jikoku;
my $code;
my $meigara;

# foreach my $stock(@STOCKS) {
my $res =  $ws->scrape(URI->new('http://quote.yahoo.co.jp/q?s=' . $Meigara));
$price = $res->{price};
$jikoku = $res->{jikoku};
print $jikoku . " " . $price . "\n";

my $i = 0;
my $j =0;
my $length;
for ( $i =5 ;$i >= 0 ; $i--){

my @list1 = split(/\//, $hajimari[$i] );
#  print $list1[0]."\n";
#  print $list1[1]."\n";
#  print $list1[2]."\n";

my @list = split(/\//, $owari[$i]);
#  print $list[0]."\n";
#  print $list[1]."\n";
#  print $list[2]."\n";
#   print $list[3]."\n";
print $hajimari[$i] . " - " . $owari[$i] . "\n";

my $scraper = scraper {
 process '//tr[@bgcolor="#ffffff"]/td[1]','sdata[]' => 'TEXT';
 process '//tr[@bgcolor="#ffffff"]/td[2]','svalue[]' => 'TEXT';
 process '//tr[@bgcolor="#ffffff"]/td[3]','lvalue[]' => 'TEXT';
 process '//tr[@bgcolor="#ffffff"]/td[4]','hvalue[]' => 'TEXT';
 process '//tr[@bgcolor="#ffffff"]/td[5]','evalue[]' => 'TEXT';
 process '//tr[@bgcolor="#ffffff"]/td[6]','mount[]' => 'TEXT';
 process '//tr[@bgcolor="#ffffff"]/td[7]','fvalue[]' => 'TEXT';
 process '//b[@class="yjXL"]','name' => 'TEXT';
};
  my $res = $scraper->scrape(URI->new("http://table.yahoo.co.jp/t?c=".$list1[0]."&a=".$list1[1]."&b=".$list1[2]."&f=".$list[0]."&d=".$list[1]."&e=".$list[2]."&g=d&s=".$Meigara.".t&y=0&z=&x=sb"));

 my @sdata = @{$res->{sdata}};
 my @svalue = @{$res->{svalue}};
 my @lvalue = @{$res->{lvalue}};
 my @hvalue = @{$res->{hvalue}};
 my @evalue = @{$res->{evalue}};
 my @mount = @{$res->{mount}};
 my @fvalue = @{$res->{fvalue}};
 my $name = $res->{name};
    $length = @svalue;

 my $hiduke = "hiduke";
  $worksheet->write( 0 , 0, $name );
  $worksheet->write( 0 , 1, encode('utf-8',$hiduke ));
  $worksheet->write( 0 , 2, "svalue");
  $worksheet->write( 0 , 3, "hvalue");
  $worksheet->write( 0 , 4, "lvalue");
  $worksheet->write( 0 , 5, "evalue");
  $worksheet->write( 0 , 6, "mount");
  $worksheet->write( 0 , 7, "fvalue");
 for ( $_= 0; $_ < $length; $_++){
#  print $sdata[$_], ",", "\"";
#  print encode('utf-8',$sdata[$_]), ",", "\"";
#  print encode('utf-8',$svalue[$_]), "\"", ",", "\"","\n";
       $svalue[$_] =~ s/\,//g;
       $lvalue[$_] =~ s/\,//g;
       $hvalue[$_] =~ s/\,//g;
       $evalue[$_] =~ s/\,//g;
       $mount[$_] =~ s/\,//g;
       $fvalue[$_] =~ s/\,//g;
  $worksheet->write($j + $length - $_ ,1, $sdata[$_]);
  $worksheet->write($j + $length - $_ ,2, $svalue[$_]);
  $worksheet->write($j + $length - $_ ,3, $lvalue[$_]);
  $worksheet->write($j + $length - $_ ,4, $hvalue[$_]);
  $worksheet->write($j + $length - $_ ,5, $evalue[$_]);
  $worksheet->write($j + $length - $_ ,6, $mount[$_]);
  $worksheet->write($j + $length - $_ ,7, $fvalue[$_]);
 }
 $j = $j + $length;
 }
 print $j . " " . $jikoku . " " . $price . "\n" ;
  $worksheet->write($j + 1 ,1, $jikoku);
  $worksheet->write($j + 1 ,5, $price);
}
close(IN);

