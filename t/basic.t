#!/usr/bin/perl

use strict;
use warnings;
use Tidy::Spreadsheet;
use Test::More;

my $spreadsheet = Tidy::Spreadsheet->new();
ok $spreadsheet->load_spreadsheet("t/data/names.xls"), 'loaded - xls';

my @headers_array = $spreadsheet->get_headers();
is_deeply \@headers_array, ['Names'], "get_headers()";

my @contents = $spreadsheet->get_contents();
is_deeply \@contents, [ ["Ashley,John,Dave,Steve"] ], "get_contents()"
  or diag explain \@contents;

my $row_one = $spreadsheet->get_row_contents(1);
is $row_one, '1:Names', "get_row_contents()";

my @search_results = $spreadsheet->row_contains('Ashley');
is_deeply \@search_results, ["2:Ashley,John,Dave,Steve"], "row_contains()"
  or diag explain \@search_results;

my @split_column = $spreadsheet->col_split( 1, ",", 3, \@headers_array );
is_deeply \@split_column, [ [ 'Ashley', 'John', 'Dave', 'Steve' ] ],
  "col_split()"
  or diag explain \@split_column;

my @split_row = $spreadsheet->row_split( 2, "," );
is_deeply \@split_row, [ ['Steve'], ['Dave'], ['John'], ['Ashley'] ],
  "row_split"
  or diag explain \@split_row;

my $edited_file = "t/data/test_split.xls";
ok $spreadsheet->save_contents( $edited_file, \@headers_array, \@split_column ),
  'content saved - xls';

ok -e $edited_file, 'file created.';

END { unlink $edited_file; }

done_testing;