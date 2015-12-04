#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/hidden-sheet.xlsx');
my $ws1 = $wb->worksheet(0);
ok(!$ws1->is_sheet_hidden(), 'Regular worksheet is not hidden');

my $ws2 = $wb->worksheet(1);
ok($ws2->is_sheet_hidden(), 'Hidden worksheet is hidden');

done_testing;
