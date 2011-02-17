#!/usr/bin/perl

use Modern::Perl;
use Tidy::Spreadsheet;
use Test::More;

my $spreadsheet = Tidy::Spreadsheet->new();

note 'Load csv file with and without delimiter supplied'; {
    is $spreadsheet->load_spreadsheet("t/data/test.csv"), 0,
      'csvload -> missing delim';

    ok $spreadsheet->load_spreadsheet( "t/data/test.csv", "," ),
      'csvload -> with delim';
}

my @csv_headers = $spreadsheet->get_headers();
my @csv_contents = $spreadsheet->get_contents(1);

note 'Check spreadsheet loaded into array correctly'; {
    is_deeply \@csv_headers,
      [
        'Heading one',
        'Heading two',
        'Heading three',
        'Heading four',
        'Heading five',
        'Heading six'
      ],
      'get_headers() - csv'
      or diag explain \@csv_headers;

    is_deeply \@csv_contents,
      [
        [ '1',         '2',     '3',        '4',   '5',   '6' ],
        [ 'a',         'b',     'c',        'd',   'e',   'f' ],
        [ 'What, the', '1 , 2', ',3',       '123', '45',  '3' ],
        [ '1',         '2',     '3',        '4',   '5',   '6' ],
        [ 'a,b',       'w',     '2,3,4,5',  '4',   '5',   '2,1,6' ],
        [ 'last,',     'test,', 'row,test', 'ok ', 'one', 'two' ]
      ],
      'get_contents() - csv'
      or diag explain \@csv_contents;
}

my @csv_splitcol = $spreadsheet->col_split( 2, ",", 2, \@csv_headers );
my @csv_rowsplit = $spreadsheet->row_split( 2, ",", 1 );

note 'Editing/splitting of the file'; {
    my @row_contains = $spreadsheet->row_contains('test');
    is_deeply \@row_contains, ['7:last,:test,:row,test:ok :one:two'],
      'row_contains() - csv'
      or diag explain \@row_contains;

    is_deeply \@csv_splitcol,
      [
        [ '1',         '2',    ' ',  ' ', '3',        '4',   '5',   '6' ],
        [ 'a',         'b',    ' ',  ' ', 'c',        'd',   'e',   'f' ],
        [ 'What, the', '1 ',   ' 2', ' ', ',3',       '123', '45',  '3' ],
        [ '1',         '2',    ' ',  ' ', '3',        '4',   '5',   '6' ],
        [ 'a,b',       'w',    ' ',  ' ', '2,3,4,5',  '4',   '5',   '2,1,6' ],
        [ 'last,',     'test', ' ',  ' ', 'row,test', 'ok ', 'one', 'two' ],
      ],
      'get_contents() - csv'
      or diag explain \@csv_splitcol;

    is_deeply \@csv_rowsplit,
      [
        [ ( 1 .. 6 ) ],
        [ ( 'a' .. 'f' ) ],
        [ ' ', ' 2' ],
        [ 'What, the', '1 ', ',3', '123', '45', '3' ],
        [ ( 1 .. 6 ) ],
        [ 'a,b',   'w',    '2,3,4,5',  '4',   '5',   '2,1,6' ],
        [ 'last,', 'test', 'row,test', 'ok ', 'one', 'two' ],
      ],
      'row_split() - csv'
      or diag explain \@csv_rowsplit;
}

my $edited_csv_file = "t/data/test_split.csv";
note 'Saving back to file.'; {
    ok $spreadsheet->save_contents( $edited_csv_file, \@csv_headers,
        \@csv_splitcol), 'file saved? - csv';

    ok -e $edited_csv_file, 'file exists - csv';
}

END { unlink $edited_csv_file; }

done_testing;

