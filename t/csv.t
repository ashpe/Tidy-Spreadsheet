#!/usr/bin/perl

use strict;
use warnings;
use Tidy::Spreadsheet;
use Test::More;

my $spreadsheet = Tidy::Spreadsheet->new();

is $spreadsheet->load_spreadsheet("t/data/test.csv"), 0,
  'csvload -> missing delim';

ok $spreadsheet->load_spreadsheet( "t/data/test.csv", "," ),
  'csvload -> with delim';

my @csv_headers = $spreadsheet->get_headers();
is_deeply \@csv_headers,
  [
    'Heading one',
    'Heading two',
    'Heading three',
    'Heading four',
    'Heading five',
    'Heading six'
  ],
  'get_headers() - csv'
  or diag explain \@csv_headers;

my @csv_contents = $spreadsheet->get_contents(1);
is_deeply \@csv_contents,
  [
    [ '1',         '2',     '3',        '4',   '5',   '6' ],
    [ 'a',         'b',     'c',        'd',   'e',   'f' ],
    [ 'What, the', '1 , 2', ',3',       '123', '45',  '3' ],
    [ '1',         '2',     '3',        '4',   '5',   '6' ],
    [ 'a,b',       'w',     '2,3,4,5',  '4',   '5',   '2,1,6' ],
    [ 'last,',     'test,', 'row,test', 'ok ', 'one', 'two' ]
  ],
  'get_contents() - csv'
  or diag explain \@csv_contents;

my @row_contains = $spreadsheet->row_contains('test');
is_deeply \@row_contains, ['7:last,:test,:row,test:ok :one:two'],
  'row_contains() - csv'
  or diag explain \@row_contains;

my @csv_splitcol = $spreadsheet->col_split( 2, ",", 2, \@csv_headers );
is_deeply \@csv_splitcol,
  [
    [ '1',         '2',    ' ',  ' ', '3',        '4',   '5',   '6' ],
    [ 'a',         'b',    ' ',  ' ', 'c',        'd',   'e',   'f' ],
    [ 'What, the', '1 ',   ' 2', ' ', ',3',       '123', '45',  '3' ],
    [ '1',         '2',    ' ',  ' ', '3',        '4',   '5',   '6' ],
    [ 'a,b',       'w',    ' ',  ' ', '2,3,4,5',  '4',   '5',   '2,1,6' ],
    [ 'last,',     'test', ' ',  ' ', 'row,test', 'ok ', 'one', 'two' ],
  ],
  'get_contents() - csv'
  or diag explain \@csv_splitcol;

my @csv_rowsplit = $spreadsheet->row_split( 2, ",", 1 );
is_deeply \@csv_rowsplit,
  [
    [ ( 1 .. 6 ) ],
    [ ( 'a' .. 'f' ) ],
    [ ' ', ' 2' ],
    [ 'What, the', '1 ', ',3', '123', '45', '3' ],
    [ ( 1 .. 6 ) ],
    [ 'a,b',   'w',    '2,3,4,5',  '4',   '5',   '2,1,6' ],
    [ 'last,', 'test', 'row,test', 'ok ', 'one', 'two' ],
  ],
  'row_split() - csv'
  or diag explain \@csv_rowsplit;

my $edited_csv_file = "t/data/test_split.csv";
ok $spreadsheet->save_contents( $edited_csv_file, \@csv_headers,
    \@csv_contents ), 'file saved? - csv';

ok -e $edited_csv_file, 'file exists - csv';

END { unlink $edited_csv_file; }

done_testing;

