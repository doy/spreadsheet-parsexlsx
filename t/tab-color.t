#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/tab-color.xlsx');
my $ws1 = $wb->worksheet(0);
is($ws1->get_tab_color, '#FF0000');

my $ws2 = $wb->worksheet(1);
is($ws2->get_tab_color, undef);

done_testing;
