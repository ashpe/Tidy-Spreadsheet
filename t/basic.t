#!/bin/usr/perl

use strict;
use warnings;
use Tidy::Spreadsheet;
use Test::More;


my $spreadsheet = Tidy::Spreadsheet->new();

ok $spreadsheet->load_spreadsheet("t/data/names.xls");


my @headers_array = $spreadsheet->get_headers();
is_deeply \@headers_array, ['Names'];

my @contents = $spreadsheet->get_contents();
is_deeply \@contents, [["Ashley,John,Dave,Steve"]] or diag explain \@contents;

done_testing;

__END__

#my $row_one = $spreadsheet->get_row_contents(1);
#my @search_results = $spreadsheet->row_contains('1');

my @edited_contents = $spreadsheet->col_split(1, ",", 3, \@headers_array);

#my @edited_contents = $spreadsheet->row_split(2, ",");

$spreadsheet->save_contents("test_split.xls", \@headers_array, \@edited_contents);

__END__

if (Tidy::Spreadsheet->load_spreadsheet($csv_file, ",") == 0) {
    print "CSV File successfully loaded\n";
}

Tidy::Spreadsheet->print_spreadsheet();
print Tidy::Spreadsheet->get_row_contents(2), "\n";

Tidy::Spreadsheet->row_contains("1");
