#!/usr/bin/env perl

use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/ahihi.xlsx');
isa_ok($wb, 'Spreadsheet::ParseExcel::Workbook');

done_testing;
