package Tidy::Spreadsheet;

use Modern::Perl;
use Moose;
use Spreadsheet::Read;
use Spreadsheet::SimpleExcel;
use Carp qw( croak );

has 'spreadsheet', is => 'rw', isa => 'ArrayRef';
has 'maxcol',      is => 'rw', isa => 'Int';
has 'maxrow',      is => 'rw', isa => 'Int';

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

sub load_spreadsheet {

    my ( $self, $file_name, $delimiter ) = @_;

    if ( $file_name =~ /.csv$/ ) {
        if ( defined $delimiter ) {
            $self->spreadsheet( ReadData( $file_name, sep => $delimiter ) );
            return 1;
        }
        else {
            return 0;
        }
    }
    elsif ( $file_name =~ /.xls$/ ) {
        $self->spreadsheet( ReadData( $file_name, parser => 'xls' ) );
        return 1;
    }
    else {
        $self->spreadsheet( ReadData($file_name) );
        return 1;
    }

}

=head2 save_contents(filename, headers, contents)

Saves file. Overwrites old file if required.

=cut

sub save_contents {

    my ( $self, $filename, $header, $content ) = @_;
    
    #TODO: Add save to CSV file as it doesn't currently work. if (csv) {} else {}

    if ($filename =~ /.csv$/) {
        unshift @{$content}, $header;
        my $temp = join "\n", map { $_ = join ",", @{$_} } @{$content};
        
        open FH, ">", $filename or die $!;
        print FH $temp;
        close FH;
        return 1;
    } 
    else {
        my $excel = Spreadsheet::SimpleExcel->new();

        $excel->add_worksheet( 'Sheet 1',
            { -headers => $header, -data => $content } );

            
        $excel->output_to_file($filename) or croak $excel->errstr();

        return 1;
    }
}

=head2 get_row_contents(row number, sheetnumber) 

Sheetnumber is optional, will default to the first spreadsheet if not set, row number is the row number!

=cut 

sub get_row_contents {
    my ( $self, $row_num, $sheet_num ) = @_;
    $sheet_num = 1 unless defined $sheet_num;

    my @get_row =
      Spreadsheet::Read::row( $self->spreadsheet->[$sheet_num], $row_num );
    my $return_value = "$row_num:" . join ':', @get_row;
    return $return_value;
}

=head2 row_contains(pattern)

Checks to see if a row contains a regular expression pattern, if matches adds to array and returns the row with row number appended to the front.

=cut

sub row_contains {

    my ( $self, $pattern ) = @_;
    my @results_array;
    my $num_sheets = @{ $self->spreadsheet } - 1;

    for my $sheet ( 1 .. $num_sheets ) {
        for my $row ( 1 .. $self->maxrow ) {
            my $cell = $self->get_row_contents($row);
            if ( $cell =~ /$pattern/ ) {
                push @results_array, $self->get_row_contents($row);
                last;
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

    #get_row_contents appends a row number to the front
    #remove by splitting and shifting
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

    my @return_contents;
    $self->maxrow( $self->spreadsheet->[$sheet]{maxrow} );
    for my $row ( 2 .. $self->maxrow ) {
        my @row_contents;
        $self->maxcol( $self->spreadsheet->[$sheet]{maxcol} );
        for my $col ( 1 .. $self->maxcol ) {
            my $cell = $self->spreadsheet->[$sheet]{ cr2cell( $col, $row ) };
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

    for my $s ( 1 .. @{$split_array} - 1 ) {
        my @tmp_array = ();
        if ( scalar( @{$split_array} ) > 2 ) {
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
        #if ( @tmp_insert) {
            @{ $insert_array->[$column][$s] } = @tmp_insert;
            #}
    }

    return 1;
}

=head2 add_insert_array(@insert_arr, @contents, $column, $self->maxcolumn); 

Adds new rows into current contents.

=cut

sub add_insert_array {
    my ( $self, $insert_array, $content_array, $row_num, $column ) = @_;

    # Splice required fields into content
    if ( $column == $self->maxcol - 1 && @{$insert_array} ) {
        foreach my $colref ( @{$insert_array} ) {
            foreach my $arr (@$colref) {
                splice @{$content_array}, $row_num, 0, $arr unless !defined($arr);
            }
        }
    }
}

=head2 row_split(row, delimiter(s), optional column) 

Splits a row into multiple rows, depending on how many are found. Can specify a specific column where the row should split, if left blank will search through every column and split all possible matches. To match based on a search string, and not specifying the row number see row_splitmatch.

=cut

sub row_split {
    my ( $self, $row, $delimiter, $col ) = @_;
    $col = 'all' unless defined $col;

    my @content = $self->get_contents();

    my @insert_array = ();
    my $row_num      = 0;

    foreach my $arrayref (@content) {
        my $column = 0;
        @insert_array = ();
        foreach my $element (@$arrayref) {
            # Removed if(defined $element) from here.
            # Check field has a value and it matches.
                if ( $element =~ /$delimiter/ && $element ne $delimiter ) {
                    my $tmp = $element;
                    my @split_array = split $delimiter, $tmp;
                    if ( $col eq 'all' ) {
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

                $self->add_insert_array( \@insert_array, \@content, $row_num,
                    $column );

                #Inc current column
                $column += 1;
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
            if ( defined $add_content[$i][$j]  ) {
                my $real_column = ( $column - 1 ) + $i;
                $content[$j][$real_column] = $add_content[$i][$j];
            }
        }
    }

    return @content;

}

1; #End!
