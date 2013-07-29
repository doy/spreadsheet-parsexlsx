#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/bug-2.xlsx');
is($wb->worksheet_count, 3);

my $ws = $wb->worksheet(0);
is($ws->get_name, 'Placement');

is_deeply([$ws->row_range], [0, 0]);
is_deeply([$ws->col_range], [0, 0]);
is_deeply($ws->{Selection}, [1, 0]);

my $cell = $ws->get_cell(0, 0);
is($cell->value, "HELLO");
is($cell->type, 'Text');

done_testing;
