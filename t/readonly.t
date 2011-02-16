#!/usr/bin/perl

use strict;
use warnings;
use Tidy::Spreadsheet;
use Test::More;
use Test::Exception;

my $readonly_file = "t/data/names_readonly.xls";

my $spreadsheet = Tidy::Spreadsheet->new();
ok $spreadsheet->load_spreadsheet($readonly_file), 'fileloaded';

my @content = $spreadsheet->get_contents();
my @headers = $spreadsheet->get_headers();

dies_ok { $spreadsheet->save_contents( $readonly_file, \@headers, 
                                        \@content ) };

dies_ok { $spreadsheet->save_contents( "t/data/test_readonly.ods",
                                        \@headers, \@content ) };

dies_ok { $spreadsheet->save_contents( "t/data/test_readonly.csv",
                                        \@headers, \@content ) };
done_testing;
