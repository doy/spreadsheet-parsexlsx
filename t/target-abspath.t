#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/target-abspath.xlsx');
is($wb->worksheet(0)->get_cell(1, 0)->value, '10213.576');

done_testing;
