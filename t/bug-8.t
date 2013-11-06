#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/bug-8.xlsx');
is($wb->worksheet_count, 3);

my $ws = $wb->worksheet(2);
my ($rmin, $rmax) = $ws->row_range;
my ($cmin, $cmax) = $ws->col_range;
is($rmin, 0);
is($rmax, -1);
is($cmin, 0);
is($cmax, -1);

done_testing;
