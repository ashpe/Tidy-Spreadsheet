package Tidy::Spreadsheet;

use warnings;
use strict;
use Spreadsheet::Read;

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

sub load_spreadsheet {

    #TODO: Add error checking once spreadsheet is loaded
    #TODO: Check if it's blank etc.
    my ($self, $file_name, $delimiter) = @_;

    if ($file_name =~ /.csv$/) {
        if (defined($delimiter)) {
            $spreadsheet = ReadData($file_name, sep => $delimiter);
            return 0;
        } else {
           die "Error; supply delimiter for csv files, check the documentation."; 
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
    my $maxrow = $spreadsheet->[1]{maxrow};
    my $maxcol = $spreadsheet->[1]{maxcol};
    my $cell = "";
    
    for(my $sheet = 1; $sheet<=$maxsheet; $sheet++) {
        for(my $row = 1; $row<=$maxrow; $row++) {
            for(my $col = 1; $col<=$maxcol; $col++) {
                $cell = $spreadsheet->[$sheet]{cr2cell($col, $row)};
                if ($cell =~ /$pattern/) {
                    $results_array[$total_results] = Tidy::Spreadsheet->get_row_contents($row);
                    print "Match found @ " . cr2cell($col, $row) . " on sheet $sheet\n";
                    $total_results += 1;
                }
                #print $cell, " ";
            }   
        }   
    }
    
    print "\nRows found;\n";
    print "@results_array\n";
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
