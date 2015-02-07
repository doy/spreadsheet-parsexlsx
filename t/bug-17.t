#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wba = Spreadsheet::ParseXLSX->new->parse('t/data/bug-17a.xlsx');
ok($wba->using_1904_date);

my $wbb = Spreadsheet::ParseXLSX->new->parse('t/data/bug-17b.xlsx');
ok(!$wbb->using_1904_date);

done_testing;
