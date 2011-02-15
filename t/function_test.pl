#!/bin/usr/perl

use strict;
use warnings;
use Tidy::Spreadsheet;

#my $csv_file = "t/test.csv";

my $spreadsheet = Tidy::Spreadsheet->new();

$spreadsheet->load_spreadsheet($ARGV[0]);

my @headers_array = $spreadsheet->get_headers();
my @contents = $spreadsheet->get_contents();

print "@headers_array\n";

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
