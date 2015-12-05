#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/target-abspath.xlsx');
like($wb->worksheet(0)->get_cell(1, 0)->value, qr/^10213\.5/);

done_testing;
