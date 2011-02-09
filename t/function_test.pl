#!/bin/usr/perl

use strict;
use warnings;
use Tidy::Spreadsheet;

#my $csv_file = "t/test.csv";

Tidy::Spreadsheet->load_spreadsheet($ARGV[0]);

my @headers_array = Tidy::Spreadsheet->get_headers();
my @contents = Tidy::Spreadsheet->get_contents();

print "@headers_array\n";
for my $i (0..$#contents) {
    for my $j (0..$#{$contents[$i]}) {
        if (defined($contents[$i][$j])) {
            print "$contents[$i][$j] ";
        }
    }
    print "\n";
}

Tidy::Spreadsheet->save_contents("test.xls", \@headers_array, \@contents);

__END__

if (Tidy::Spreadsheet->load_spreadsheet($csv_file, ",") == 0) {
    print "CSV File successfully loaded\n";
}

Tidy::Spreadsheet->print_spreadsheet();
print Tidy::Spreadsheet->get_row_contents(2), "\n";

Tidy::Spreadsheet->row_contains("1");
