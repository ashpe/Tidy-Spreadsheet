#!/bin/usr/perl

use strict;
use warnings;
use Tidy::Spreadsheet;
use Test::More;

my $spreadsheet = Tidy::Spreadsheet->new();
is $spreadsheet->load_spreadsheet("t/data/test.csv"), 0;
ok $spreadsheet->load_spreadsheet("t/data/test.csv", ",");

done_testing;

