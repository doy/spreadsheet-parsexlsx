#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/hidden-row-and-column.xlsx');
my $ws = $wb->worksheet(0);

ok(!$ws->is_row_hidden(0), 'Regular row is not hidden');
ok( $ws->is_row_hidden(1), 'Hidden row is hidden');

ok(!$ws->is_col_hidden(0), 'Regular column is not hidden');
ok( $ws->is_col_hidden(1), 'Hidden column is hidden');

done_testing;
