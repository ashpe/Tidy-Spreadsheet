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

=head1 EXPORT

A list of functions that can be exported.  You can delete this section
if you don't export anything, such as for a purely object-oriented module.

=head1 SUBROUTINES/METHODS

=head2 function1

=cut

sub load_spreadsheet {

    #TODO: Add error checking once spreadsheet is loaded

    my ($self, $file_name, $delimiter) = @_;

    if ($file_name =~ /.csv$/) {
        if (!$delimiter) {
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
        print $spreadsheet->[1]{cr2cell(1,1)}, "\n";
        print "@row";
     }
}

sub get_row_contents {
    my ($self, $row_num) = @_;

    my @get_row = Spreadsheet::Read::cellrow($spreadsheet->[1], $row_num);
    my $return_value = "$row_num:" . join(":", @get_row);
    print "$return_value\n";
}

=head2 function2

=cut

sub function2 {
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
