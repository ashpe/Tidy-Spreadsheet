#!perl -T

use Test::More tests => 1;

BEGIN {
    use_ok( 'Tidy::Spreadsheet' ) || print "Bail out!
";
}

diag( "Testing Tidy::Spreadsheet $Tidy::Spreadsheet::VERSION, Perl $], $^X" );
