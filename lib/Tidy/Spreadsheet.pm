package Tidy::Spreadsheet;

use warnings;
use strict;
use Readonly;
use Spreadsheet::Read;
use Spreadsheet::SimpleExcel;
use Carp qw( croak );
Readonly my $NOT_PROVIDED => -1;

=head1 NAME

Tidy::Spreadsheet - The great new Tidy::Spreadsheet!

=head1 VERSION

Version 0.01

=cut

our $VERSION = '0.01';

=head1 SYNOPSIS

Quick summary of what the module does.

Perhaps a little code snippet.

    use Tidy::Spreadsheet;

    my $foo = Tidy::Spreadsheet->new();
    ...

=head1 SUBROUTINES/METHODS

=head2 load_spreadsheet

Load a spreadsheet into the modules for use with other functions.

example; Tidy::Spreadsheet->load_spreadsheet("spreadsheet.xls");

or for a csv file, you need to specify the delimiters.

example; Tidy::Spreadsheet->load_spreadsheet("spreadsheet.csv", ","); 

=cut

sub new {
    my $package = shift;
    return bless {}, $package;
}

sub load_spreadsheet {

    #TODO: Add error checking once spreadsheet is loaded
    #TODO: Check if it's blank etc.
    my ( $self, $file_name, $delimiter ) = @_;

    if ( $file_name =~ /.csv$/ ) {
        if ( defined $delimiter ) {
            $self->{spreadsheet} = ReadData( $file_name, sep => $delimiter );
            return 1;
        }
        else {
            return 0;
        }
    }
    elsif ( $file_name =~ /.xls$/ ) {
        $self->{spreadsheet} = ReadData( $file_name, parser => 'xls' );
        return 1;
    }
    else {
        $self->{spreadsheet} = ReadData($file_name);
        return 1;
    }

}

=head2 save_contents(filename, headers, contents)

Saves file. Overwrites old file if required.

=cut

sub save_contents {

    my ( $self, $filename, $header, $content ) = @_;

    my $excel = Spreadsheet::SimpleExcel->new();

    $excel->add_worksheet( 'Sheet 1',
        { -headers => $header, -data => $content } );

    $excel->output_to_file($filename) or croak $excel->errstr();
    
    return 1;
}

=head2 get_row_contents(row number, sheetnumber) 

Sheetnumber is optional, will default to the first spreadsheet if not set, row number is the row number!

Further explination on sheetnumber, its the current spreadsheet tab number you are viewing starting at 1.

=cut 

sub get_row_contents {
    my ( $self, $row_num, $sheet_num ) = @_;
    $sheet_num = 1 unless defined $sheet_num;

    my @get_row =
      Spreadsheet::Read::row( $self->{spreadsheet}->[$sheet_num], $row_num );
    my $return_value = "$row_num:" . join ':', @get_row;
    return $return_value;
}

=head2 row_contains(pattern)

Checks to see if a row contains a regular expression pattern, if matches adds to array and returns the row with row number appended to the front.

=cut

sub row_contains {

    my ( $self, $pattern ) = @_;

    my @results_array;
    my $total_results = 0;
    my $maxsheet      = $self->{spreadsheet}->[0]{sheets};
    my $cell          = ' ';

    for my $sheet ( 1 .. $maxsheet ) {
        my $maxrow = $self->{spreadsheet}->[$sheet]{maxrow};
        my $maxcol = $self->{spreadsheet}->[$sheet]{maxcol};
        for my $row ( 2 .. $maxrow ) {
            for my $col ( 1 .. $maxcol ) {
                $cell = $self->{spreadsheet}->[$sheet]{ cr2cell( $col, $row ) };
                if ( $cell =~ /$pattern/ ) {
                    $results_array[$total_results] =
                      $self->get_row_contents($row);
                    $total_results += 1;
                }
            }
        }
    }
    return @results_array;
}

=head2 get_headers() 

Returns array of headers.

=cut

sub get_headers {
    my ($self) = @_;

    my @return_array = split ':', $self->get_row_contents(1);
    shift @return_array;
    return @return_array;
}

=head2 get_contents(Sheetnumber)

    Returns all the content of the sheet provided for currently loaded spreadsheet. Removes header from return value.

=cut

sub get_contents {

    my ( $self, $sheet ) = @_;
    $sheet = 1 unless defined $sheet;

    my $maxrow = $self->{spreadsheet}->[$sheet]{maxrow};
    my $maxcol = $self->{spreadsheet}->[$sheet]{maxcol};
    my $cell   = ' ';
    my @return_contents;

    for my $row ( 2 .. $maxrow ) {
        my @row_contents;
        for my $col ( 1 .. $maxcol ) {
            $cell = $self->{spreadsheet}->[$sheet]{ cr2cell( $col, $row ) };
            push @row_contents, $cell;
        }
        push @return_contents, \@row_contents;
    }

    return @return_contents;

}

=head2 prepare_to_insert(@splitarray, @array to prepare into) 

Prepares the array for inserting back into the excel file.

=cut

sub prepare_to_insert {

    my ( $self, $column, $split_array, $insert_array ) = @_;

    for ( my $s = 1 ; $s < scalar @{$split_array} ; $s++ ) {
        my @tmp_array = ();
        if ( scalar( @{$split_array} ) >= 2 ) {
            @tmp_array = $split_array->[$s];
        }
        else {
            $tmp_array[0] = $split_array->[1];
        }

        my @tmp_insert = ();
        for ( 0 .. $column ) {
            if ( $_ != $column ) {
                $tmp_insert[$_] = ' ';
            }
            else {
                $tmp_insert[$_] = $tmp_array[0];
            }
        }
        if (@tmp_insert) {
            @{ $insert_array->[$column][$s] } = @tmp_insert;
        }
    }
    return;
}

=head2 row_split(row, delimiter(s), optional column) 

Splits a row into multiple rows, depending on how many are found. Can specify a specific column where the row should split, if left blank will search through every column and split all possible matches. To match based on a search string, and not specifying the row number see row_splitmatch.

=cut

sub row_split {
    my ( $self, $row, $delimiter, $col ) = @_;
    $col = $NOT_PROVIDED unless defined $col;

    my @content = $self->get_contents();

    my @insert_array = ();
    my $row_num      = 0;
    my $maxcol       = $self->{spreadsheet}->[1]{maxcol};

    foreach my $arrayref (@content) {
        my $column = 0;
        @insert_array = ();
        foreach my $element (@$arrayref) {

            # Check field has a value and it matches.
            if ( defined $element ) {
                if ( $element =~ /$delimiter/ && $element ne $delimiter ) {
                    my $tmp = $element;
                    my @split_array = split $delimiter, $tmp;
                    if ( $col == $NOT_PROVIDED ) {

                        #Overwrite our old value/set new intoarray
                        $element = $split_array[0];
                        $self->prepare_to_insert( $column, \@split_array,
                            \@insert_array );
                    }
                    elsif ( $col == $column ) {

                        # Do the same but only on specified column
                        $element = $split_array[0];
                        $self->prepare_to_insert( $column, \@split_array,
                            \@insert_array );
                    }
                }

                # Splice required fields into content
                if ( $column == $maxcol - 1 && @insert_array ) {
                    foreach my $colref (@insert_array) {
                        foreach my $arr (@$colref) {
                            if ( defined $arr ) {
                                splice @content, $row_num, 0, $arr;
                            }
                        }
                    }
                }

                #Inc current column
                $column += 1;
            }
        }

        #Inc current row
        $row_num += 1;
    }

    return @content;
}

=head2 add_columns_to(at_column, amount of new columns, data to add to, headers reference)

Used with col split to add blank columns into our content. Requires
reference to our headers so all new data will be aligned.

=cut

sub add_columns_to {

    my ( $self, $column, $new_columns, $content, $headers_arr ) = @_;

    foreach my $row ( @{$content} ) {
        for my $i ( 0 .. $new_columns - 1 ) {
            if ( !$row->[ $column + $i ] ) {
                $row->[ $column + $i ] = ' ';
            }
            else {
                splice @{$row}, $column + $i, 0, ' ';
            }
        }
    }

    for my $j ( 0 .. $new_columns - 1 ) {
        if ( !$headers_arr->[ $column + $j ] ) {
            $headers_arr->[ $column + $j ] = ' ';
        }
        else {
            splice @{$headers_arr}, $column + $j, 0, ' ';
        }
    }

}

=head2 get_new_columns( @ content , $column, $new column)

Returns the data from $column, split into the required fields for adding

=cut

sub get_new_columns {

    my ( $self, $content, $column, $new_column, $delimiter ) = @_;
    my $row_content  = '';
    my $content_size = @{$content} - 1;
    my @arr_to_add   = qw{};

    foreach my $row ( 0 .. $content_size ) {
        $row_content = $content->[$row]->[ $column - 1 ];
        if ( $row_content =~ /$delimiter/ && defined($row_content) ) {
            my @split_arr      = split $delimiter, $row_content;
            my $split_arr_size = @split_arr - 1;

            for my $i ( 0 .. $split_arr_size ) {
                $arr_to_add[$i][$row] = $split_arr[$i];
            }

        }
    }

    return @arr_to_add;
}

=head2 col_split(column, delimiter, number_of_new_columns, ref headers)

Splits a value in one column, into multiple columns based on the delimiter. Requires a reference to our headers array, so we can alter them based on new columns.

=cut

sub col_split {

    my ( $self, $column, $delimiter, $new_columns, $headers_arr ) = @_;
    my @content     = $self->get_contents();
    my @add_content =
      $self->get_new_columns( \@content, $column, $new_columns, $delimiter );

    $self->add_columns_to( $column, $new_columns, \@content, $headers_arr );

    for my $i ( 0 .. @add_content - 1 ) {
        for my $j ( 0 .. @{ $add_content[$i] } - 1 ) {
            if ( defined( $add_content[$i][$j] ) ) {
                my $real_column = ( $column - 1 ) + $i;
                $content[$j][$real_column] = $add_content[$i][$j];
            }
        }
    }

    return @content;

}

=head1 AUTHOR

Ashley Pope, C<< <ashleyp at cpan.org> >>

=head1 BUGS

Please report any bugs or feature requests to C<bug-tidy-spreadsheet at rt.cpan.org>, or through
the web interface at L<http://rt.cpan.org/NoAuth/ReportBug.html?Queue=Tidy-Spreadsheet>.  I will be notified, and then you'll
automatically be notified of progress on your bug as I make changes.

=head1 SUPPORT

You can find documentation for this module with the perldoc command.

    perldoc Tidy::Spreadsheet


You can also look for information at:

=over 4

=item * RT: CPAN's request tracker

L<http://rt.cpan.org/NoAuth/Bugs.html?Dist=Tidy-Spreadsheet>

=item * AnnoCPAN: Annotated CPAN documentation

L<http://annocpan.org/dist/Tidy-Spreadsheet>

=item * CPAN Ratings

L<http://cpanratings.perl.org/d/Tidy-Spreadsheet>

=item * Search CPAN

L<http://search.cpan.org/dist/Tidy-Spreadsheet/>

=back


=head1 ACKNOWLEDGEMENTS


=head1 LICENSE AND COPYRIGHT

Copyright 2011 Ashley Pope.

This program is free software; you can redistribute it and/or modify it
under the terms of either: the GNU General Public License as published
by the Free Software Foundation; or the Artistic License.

See http://dev.perl.org/licenses/ for more information.


=cut

1;    # End of Tidy::Spreadsheet
