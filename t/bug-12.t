#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/bug-12.xlsx');
is($wb->worksheet_count, 1);

my $ws = $wb->worksheet(0);
my ($rmin, $rmax) = $ws->row_range;
my ($cmin, $cmax) = $ws->col_range;
is($rmin, 0);
is($rmax, 0);
is($cmin, 0);
is($cmax, 3);

is($ws->get_cell(0, 0)->value, 7);
is($ws->get_cell(0, 1)->value, 3);
is($ws->get_cell(0, 2)->value, 30);
is($ws->get_cell(0, 3)->value, 'Kuku');

done_testing;
