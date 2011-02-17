#!/usr/bin/perl

use Modern::Perl;
use Tidy::Spreadsheet;
use Test::More;

my $spreadsheet = Tidy::Spreadsheet->new();
ok $spreadsheet->load_spreadsheet("t/data/test.ods"), 'load ods file';

note 'Retrieve/editing row contents';
{
    my $row_contents = $spreadsheet->get_row_contents( 1, 1 );
    is $row_contents,
"1:Heading one:Heading two:Heading three:Heading four:Heading five:Heading six",
      "row_contents with sheet specified";

    my @row_split = $spreadsheet->row_split( 2, "," );
    is_deeply \@row_split,
      [
        [ ( 1 .. 6 ) ],
        [ ( 'a' .. 'f' ) ],
        [ ' ', ' ',   '3' ],
        [ ' ', ' 2' ],
        [' the'],
        [ 'What', '1 ', '', '123', '45', '3' ],
        [ ( 1 .. 6 ) ],
        [ ' ', ' ', ' ',  ' ', ' ', '6' ],
        [ ' ', ' ', ' ',  ' ', ' ', '1' ],
        [ ' ', ' ', '5' ],
        [ ' ', ' ', '4' ],
        [ ' ', ' ', '3' ],
        ['b'],
        [ 'a',    'w',    '2',     '4',   '5',   '2' ],
        [ ' ',    ' ',    'test' ],
        [ 'last', 'test', 'row',   'ok ', 'one', 'two' ],
      ],
      'row_split - ods'
      or diag explain \@row_split;
}

done_testing;
