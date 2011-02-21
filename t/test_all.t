#!/usr/bin/perl

use Modern::Perl;
use Tidy::Spreadsheet;
use Test::More;


my @file_types = qw/test.csv test.xls test.ods/;
my @rem_files = ();

foreach my $next_file (@file_types) {

    my $spreadsheet = Tidy::Spreadsheet->new();
    my $actual_file = "t/data/" . $next_file;

    note "Load file: $next_file"; {
        
        if ($actual_file =~ /.csv/) {
            is $spreadsheet->load_spreadsheet($actual_file), 0,
              "load file: $next_file (no delim)";

            ok $spreadsheet->load_spreadsheet( $actual_file, "," ),
              "load file: $next_file (delim)";
        } else {
            ok $spreadsheet->load_spreadsheet($actual_file),
               "load file: $next_file";
        } 
    }

    my @csv_headers = $spreadsheet->get_headers();
    my @csv_contents = $spreadsheet->get_contents(1);

    note "Check $next_file loaded into array correctly"; {
        is_deeply \@csv_headers,
          [
            'Heading one',
            'Heading two',
            'Heading three',
            'Heading four',
            'Heading five',
            'Heading six'
          ],
          "get_headers() - $next_file"
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
          "get_contents() - $next_file"
          or diag explain \@csv_contents;
    }

    my @csv_splitcol = $spreadsheet->col_split( 2, ",", 2, \@csv_headers );
    my @csv_rowsplit = $spreadsheet->row_split( 2, ",", 1 );

    note "Editing/splitting $next_file"; {
        my @row_contains = $spreadsheet->row_contains('test');
        is_deeply \@row_contains, ['last,:test,:row,test:ok :one:two'],
          "row_contains() - $next_file"
          or diag explain \@row_contains;

		my @row_split2 = $spreadsheet->row_split( 2, "," );
		is_deeply \@row_split2,
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
                        "row_split2(): $next_file"
                        or diag explain \@row_split2;
            

        is_deeply \@csv_splitcol,
          [
            [ '1',         '2',    ' ',  ' ', '3',        '4',   '5',   '6' ],
            [ 'a',         'b',    ' ',  ' ', 'c',        'd',   'e',   'f' ],
            [ 'What, the', '1 ',   ' 2', ' ', ',3',       '123', '45',  '3' ],
            [ '1',         '2',    ' ',  ' ', '3',        '4',   '5',   '6' ],
            [ 'a,b',       'w',    ' ',  ' ', '2,3,4,5',  '4',   '5',   '2,1,6' ],
            [ 'last,',     'test', ' ',  ' ', 'row,test', 'ok ', 'one', 'two' ],
          ],
          "get_contents() - $next_file"
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
          "row_split() - $next_file"
          or diag explain \@csv_rowsplit;
    }

    my $edited_csv_file = $actual_file . ".tmp";
    note "Saving $next_file back to file."; {
        ok $spreadsheet->save_contents( $edited_csv_file, \@csv_headers,
            \@csv_splitcol), "file saved? - $next_file";

        ok -e $edited_csv_file, "file exists - $next_file";
    }
    push @rem_files, $edited_csv_file;
}

END { foreach my $rem (@rem_files) { unlink $rem; } }

done_testing;

