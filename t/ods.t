#!/bin/usr/perl

use strict;
use warnings;
use Tidy::Spreadsheet;
use Test::More;

my $spreadsheet = Tidy::Spreadsheet->new();
ok $spreadsheet->load_spreadsheet("t/data/test.ods");

my @row_split = $spreadsheet->row_split(2, ",", 1);

is_deeply \@row_split, [['1','2','3','4','5','6'],['a','b','c','d','e','f'],[' ',' 2'],['What, the','1 ',',3','123','45','3'],['1','2','3','4','5','6'],['a,b','w','2,3,4,5','4','5','2,1,6'],['last,','test','row,test','ok ','one','two']] or diag explain \@row_split;



done_testing;
