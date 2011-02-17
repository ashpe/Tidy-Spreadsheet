#!/usr/bin/perl

use Modern::Perl;
use Tidy::Spreadsheet;
use Test::Most;


my $readonly_file = "t/data/names_readonly.xls";

my $spreadsheet = Tidy::Spreadsheet->new();
ok $spreadsheet->load_spreadsheet($readonly_file), 'fileloaded';

my @content = $spreadsheet->get_contents();
my @headers = $spreadsheet->get_headers();


note 'attempting to overwrite read only file'; {

    dies_ok { $spreadsheet->save_contents( $readonly_file, \@headers, 
                                            \@content ) };

    dies_ok { $spreadsheet->save_contents( "t/data/test_readonly.ods",
                                            \@headers, \@content ) };

    dies_ok { $spreadsheet->save_contents( "t/data/test_readonly.csv",
                                            \@headers, \@content ) };
}

done_testing;
