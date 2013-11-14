#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/bug-13.xlsx');
is($wb->get_filename, 't/data/bug-13.xlsx');

done_testing;
