package Tidy::Spreadsheet;

use warnings;
use strict;
use Spreadsheet::Read;
use Spreadsheet::SimpleExcel;
use Carp qw( croak );
my $spreadsheet = "";

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
    return bless({}, $package);
}

sub load_spreadsheet {
    #TODO: Add error checking once spreadsheet is loaded
    #TODO: Check if it's blank etc.
    my ($self, $file_name, $delimiter) = @_;

    if ($file_name =~ /.csv$/) {
        if (defined($delimiter)) {
            $spreadsheet = ReadData($file_name, sep => $delimiter);
            return 0;
        } else {
            croak "Error; no delimiter supplied for csv file"; 
        }
    }
    elsif ($file_name =~ /.xls$/) {
        $spreadsheet = ReadData($file_name, parser => "xls");
        return 0;
    }
    else {
        $spreadsheet = ReadData($file_name);
        return 0;
    }

}

sub print_spreadsheet() {
    if (!$spreadsheet) {
        print "No spreadsheet has been loaded.\n";
    } else {
        my @row = Spreadsheet::Read::row($spreadsheet->[1],1);
        print "@row\n";
    }
    return;
}

=head2 save_contents(filename, headers, contents)

Saves file. Overwrites old file if required.

=cut

sub save_contents {

    my ($self, $filename, $header, $content) = @_;

    my $excel = Spreadsheet::SimpleExcel->new();

    $excel->add_worksheet('Sheet 1', {-headers => $header, -data => $content});

    $excel->output_to_file($filename) or die $excel->errstr();
}


=head2 get_row_contents(row number, sheetnumber) 

Sheetnumber is optional, will default to the first spreadsheet if not set, row number is the row number!

Further explination on sheetnumber, its the current spreadsheet tab number you are viewing starting at 1.

=cut 

sub get_row_contents {
    my ($self, $row_num, $sheet_num) = @_;
    $sheet_num = 1 unless defined($sheet_num); 

    my @get_row = Spreadsheet::Read::row($spreadsheet->[$sheet_num], $row_num);
    my $return_value = "$row_num:" . join(":", @get_row);
    return $return_value;
}

=head2 row_contains(pattern)

Checks to see if a row contains a regular expression pattern, if matches adds to array and returns the row with row number appended to the front.

=cut

sub row_contains {

    my ($self, $pattern) = @_;

    my @results_array;
    my $total_results = 0;
    my $maxsheet = $spreadsheet->[0]{sheets};    
    my $maxrow = 0;
    my $maxcol = 0;
    my $cell = "";

    for(my $sheet = 1; $sheet<=$maxsheet; $sheet++) {
        $maxrow = $spreadsheet->[$sheet]{maxrow};
        $maxcol = $spreadsheet->[$sheet]{maxcol};

        for(my $row = 2; $row<=$maxrow; $row++) {
            for(my $col = 1; $col<=$maxcol; $col++) {
                $cell = $spreadsheet->[$sheet]{cr2cell($col, $row)};
                if ($cell =~ /$pattern/) {
                    $results_array[$total_results] = $self->get_row_contents($row);
                    print "Match found @ " . cr2cell($col, $row) . " on sheet $sheet\n";
                    $total_results += 1;
                }
                #print $cell, " ";
            }   
        }   
    }

    print "\nRows found;\n";
    print "@results_array\n";
    return @results_array;
}

=head2 get_headers() 

Returns array of headers.

=cut

sub get_headers {
    my ($self) = @_;

    my @return_array = split(":", $self->get_row_contents(1));
    shift @return_array;
    return @return_array;
}

=head2 get_contents(Sheetnumber)

    Returns all the content of the sheet provided for currently loaded spreadsheet. Removes header from return value.

=cut

sub get_contents {

    my ($self, $sheet) = @_;
    $sheet = 1 unless defined($sheet);

    my $maxrow=$spreadsheet->[$sheet]{maxrow};
    my $maxcol=$spreadsheet->[$sheet]{maxcol};
    my $cell = "";
    my @return_contents;

    for(my $row=2; $row<=$maxrow; $row++) {
        my @row_contents;
        for(my $col=1; $col<=$maxcol; $col++) {
            $cell = $spreadsheet->[$sheet]{cr2cell($col,$row)};            push(@row_contents, $cell);
        }
        push(@return_contents, \@row_contents);
    } 

    return @return_contents;    

}


=head2 row_split(row, delimiter(s), optional column) 

Splits a row into multiple rows, depending on how many are found. Can specify a specific column where the row should split, if left blank will search through every column and split all possible matches. To match based on a search string, and not specifying the row number see row_splitmatch.

=cut

sub row_split {
    my ($self, $row, $delimiter, $col) = @_;
    $col = -1 unless defined($col);

    my @content = $self->get_contents();

    my @insert_array = ();
    my $row_num = 0;
    my $maxcol = $spreadsheet->[1]{maxcol};

    foreach my $arrayref (@content) {
        my $column = 0;
        @insert_array = ();
        foreach my $element (@$arrayref) {
            # Check field has a value and it matches.
            if (defined($element)) {
                if ($element =~ /$delimiter/ && $element ne $delimiter) {
                    my $tmp = $element;
                    my @split_array = split($delimiter, $tmp);

                    if ($col == -1) {

                        #Overwrite our old value/set new intoarray
                        $element = $split_array[0];
                        for(my $s=1; $s <= $#split_array; $s++) {
                            my @tmp_array = ();                  
                            if ($#split_array >= 2) {
                                @tmp_array = $split_array[$s];
                            } else {
                                $tmp_array[0] = $split_array[1];
                            }
                            my @tmp_insert = ();
                            for(my $l = 0;$l <= $column; $l++) {

                                if ($l != $column) {
                                    $tmp_insert[$l]= " ";
                                }
                                else {
                                    $tmp_insert[$l]=$tmp_array[0];
                                } 
                            }
                            if (@tmp_insert) {
                                @{$insert_array[$column][$s]}=@tmp_insert;
                            }
                        }
                    }
                    elsif ($col == $column) {
                        # Do the same but only on specified column    
                        $element = $split_array[0];
                        $insert_array[$column] = $split_array[1];
                    }
                }
                # Splice required fields into content
                if ($column == $maxcol-1 && @insert_array) {
                    foreach my $colref (@insert_array) {
                        foreach my $arr (@$colref) {
                            if (defined($arr)) {
                                splice @content, $row_num, 0, $arr;
                            }
                        }
                    }

                }
                #Inc current column
                $column +=1;
            }
        }
        #Inc current row
        $row_num += 1;
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

1; # End of Tidy::Spreadsheet
