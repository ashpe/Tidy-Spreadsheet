#!/bin/usr/perl

use strict;
use warnings;
use Tidy::Spreadsheet;

my $csv_file = "t/test.csv";

if (Tidy::Spreadsheet->load_spreadsheet($csv_file) == 0) {
    print "CSV File successfully loaded\n";
}

Tidy::Spreadsheet->print_spreadsheet();
Tidy::Spreadsheet->get_row_contents(1);

Tidy::Spreadsheet->row_contains("Hey");
