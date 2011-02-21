#!/usr/bin/perl

use Modern::Perl;
use Tidy::Spreadsheet;
use Test::Most;
use Smart::Comments;

my @readonly_files = qw/test_readonly.ods
                        test_readonly.csv
                        names_readonly.xls/;

my $readonly_dir = "t/data/";

my $spreadsheet = Tidy::Spreadsheet->new();
ok $spreadsheet->load_spreadsheet($readonly_dir . $readonly_files[0]), 
                                  'fileloaded';

my @content = $spreadsheet->get_contents();
my @headers = $spreadsheet->get_headers();

note 'Attempting to overwrite read only files'; {

    foreach  (@readonly_files) { 
        dies_ok { $spreadsheet->save_contents($readonly_dir . $_, \@headers, 
                  \@content ) }  "overwrite readonly: $_";
    }

}

done_testing;
